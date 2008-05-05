#---------------------------------------------------------------------
# $Header: /Perl/OlleDB/t/B_filestream.t 5     08-05-04 18:47 Sommar $
#
# Tests for OpenSqlFilestream.
#
# $History: B_filestream.t $
# 
# *****************  Version 5  *****************
# User: Sommar       Date: 08-05-04   Time: 18:47
# Updated in $/Perl/OlleDB/t
# Don't run the test without SQLNCLI10.
#
# *****************  Version 4  *****************
# User: Sommar       Date: 08-05-02   Time: 0:44
# Updated in $/Perl/OlleDB/t
# Changed the check for whether FILESTREAM is enabled.
#
# *****************  Version 3  *****************
# User: Sommar       Date: 08-02-17   Time: 18:01
# Updated in $/Perl/OlleDB/t
# Added allocation length to the last call to OpenSqlFilestream.
#
# *****************  Version 2  *****************
# User: Sommar       Date: 07-12-02   Time: 21:41
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 1  *****************
# User: Sommar       Date: 07-11-26   Time: 22:45
# Created in $/Perl/OlleDB/t
#---------------------------------------------------------------------

use strict;
use Win32::SqlServer qw(:DEFAULT :consts);
use File::Basename qw(dirname);
use Win32API::File;

require &dirname($0) . '\testsqllogin.pl';


$^W = 1;
$| = 1;

my $X = testsqllogin();
my ($sqlver) = split(/\./, $X->{SQL_version});

if ($sqlver < 10) {
   print "1..0 # Skipped: FileStream not available on SQL 2005 and earlier.\n";
   exit;
}

if ($X->{Provider} < PROVIDER_SQLNCLI10) {
   print "1..0 # Skipped: Need SQL Server Client 10 to use Filestream.\n";
   exit;
}


my $fs_config = sql_one(<<SQLEND, SCALAR);
SELECT value_in_use FROM sys.configurations WHERE name = 'filestream access level'
SQLEND
if ($fs_config < 2) {
   print "1..0 # Skipped: Instance not configured for remote filestream access.\n";
   exit;
}

my $username = sql_one("SELECT SYSTEM_USER", SCALAR);
if ($username !~ /\\/) {
   print "1..0 # Skipped: filestream requires Windows authentication.\n";
   exit;
}

print "1..8\n";

# Create a test database with a filestream filegroup.
$X->sql(<<'SQLEND');
CREATE DATABASE Olle$DB
ALTER DATABASE Olle$DB ADD FILEGROUP fs CONTAINS FILESTREAM
SQLEND

# We need to know path for data file to determine where to create the
# filestream container.
my $dbpath = $X->sql_one(<<'SQLEND', SCALAR);
SELECT physical_name
FROM   Olle$DB.sys.database_files
WHERE  file_id = 1
SQLEND
$dbpath =~ s/\.mdf$//;
$dbpath .= ".datadir";

# Now we can add the file group.
$X->sql(<<SQLEND);
ALTER DATABASE Olle\$DB ADD FILE
    (NAME = 'fs', FILENAME = '$dbpath') TO FILEGROUP fs
SQLEND

# This is our test strings.
my $yksi  = "Somliga säger att somliga somrar somnar somliga.\n" x 2000;
my $kaksi = "Nelly Nilsson nöjer sig numera näppeligen med nio nötter till natten.\n" x 2000;
my $kolme = "Handlar Hansons halta höna har haft hosta hela halva hösten.\n" x 1000;
my $negy  = "Elva elaka elefanter erövrade Enköping\n";

# Move to the database and create a table with three rows in it.
$X->sql('USE Olle$DB');
$X->sql(<<'SQLEND', {'@yksi' => ['varchar', $yksi], '@kolme' => ['varchar', $kolme]});
CREATE TABLE fstest (guid uniqueidentifier          NOT NULL ROWGUIDCOL UNIQUE,
                     name varchar(23)               NOT NULL PRIMARY KEY,
                     data varbinary(MAX) FILESTREAM NULL)

INSERT fstest (guid, name, data)
   VALUES(newid(), 'Yksi', cast(@yksi AS varbinary(MAX))),
         (newid(), 'Kaksi', 0x),
         (newid(), 'Kolme', cast(@kolme AS varbinary(MAX)))

SQLEND

# We're all set for testing. Let's try reading data.
my ($path, $context, $fh, $buffer, $ret);
$X->{BinaryAsStr} = 1;
($path, $context) = $X->sql(<<SQLEND, LIST, SINGLEROW);
BEGIN TRANSACTION
SELECT data.PathName(), get_filestream_transaction_context()
FROM   fstest
WHERE  name = 'Yksi'
SQLEND

$fh = $X->OpenSqlFilestream($path, FILESTREAM_READ, $context);
if ($fh > 0) {
   print "ok 1\n";
}
else {
   print "not ok 1\n";
}

$ret = Win32API::File::ReadFile($fh, $buffer, 200000, [], []);
if ($ret) {
   print "ok 2\n";
}
else {
   print "not ok 2 # ReadFile failed with $^E\n";
}

if ($buffer eq $yksi) {
   print "ok 3\n";
}
else {
   print "not ok 3\n";
}

# Close this transaction.
Win32API::File::CloseHandle($fh);
$X->sql('ROLLBACK TRANSACTION');

# Try writing.
$X->{BinaryAsStr} = 0;
($path, $context) = $X->sql(<<SQLEND, LIST, SINGLEROW);
BEGIN TRANSACTION
SELECT data.PathName(), get_filestream_transaction_context()
FROM   fstest
WHERE  name = 'Kaksi'
SQLEND

# The option is just to test that options work.
$fh = $X->OpenSqlFilestream($path, FILESTREAM_WRITE, $context,
                            SQL_FILESTREAM_OPEN_FLAG_NO_WRITE_THROUGH);
if ($fh > 0) {
   print "ok 4\n";
}
else {
   print "not ok 4\n";
}

$ret = Win32API::File::WriteFile($fh, $kaksi, 0, [], []);
if ($ret) {
   print "ok 5\n";
}
else {
   print "not ok 5 # WriteFile failed with $^E\n";
}
# Close this transaction.
Win32API::File::CloseHandle($fh);
$X->sql('COMMIT TRANSACTION');

# And check the data.
$buffer = $X->sql_one(<<SQLEND, SCALAR);
SELECT convert(varchar(MAX), data)
FROM   fstest
WHERE  name = 'Kaksi'
SQLEND
if ($buffer eq $kaksi) {
   print "ok 6\n";
}
else {
   print "not ok 6\n";
}


$X->{BinaryAsStr} = 'x';
($path, $context) = $X->sql(<<SQLEND, LIST, SINGLEROW);
BEGIN TRANSACTION
SELECT data.PathName(), get_filestream_transaction_context()
FROM   fstest
WHERE  name = 'Kolme'
SQLEND

$fh = $X->OpenSqlFilestream($path, FILESTREAM_READWRITE, $context,
                            SQL_FILESTREAM_OPEN_FLAG_RANDOM_ACCESS,
                            {High => 2, Low => 1000000});
if ($fh > 0) {
   print "ok 7\n";
}
else {
   print "not ok 7\n";
}


# Close this transaction.
Win32API::File::CloseHandle($fh);
$X->sql('COMMIT TRANSACTION');

undef $buffer;

# And check the data.
$buffer = $X->sql_one(<<SQLEND, SCALAR);
SELECT convert(varchar(MAX), data)
FROM   fstest
WHERE  name = 'Kolme'
SQLEND
if ($buffer eq '') {
   print "ok 8\n";
}
else {
   print "not ok 8\n";
}



$X->sql('USE master');
$X->sql('DROP DATABASE Olle$DB');

exit;

