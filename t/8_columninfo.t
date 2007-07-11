#---------------------------------------------------------------------
# $Header: /Perl/OlleDB/t/8_columnsinfo.t 2     07-07-07 22:26 Sommar $
#
# This test suite tests the data type information returned by
# getcolumninfo.
#
# $History: 8_columnsinfo.t $
# 
# *****************  Version 2  *****************
# User: Sommar       Date: 07-07-07   Time: 22:26
# Updated in $/Perl/OlleDB/t
# Added checks so that we don't run XML and UDT with SQLOLEDB or CLR
# disabled.
#
# *****************  Version 1  *****************
# User: Sommar       Date: 07-07-07   Time: 16:46
# Created in $/Perl/OlleDB/t
#---------------------------------------------------------------------

use strict;
use Win32::SqlServer qw(:DEFAULT :consts);
use File::Basename qw(dirname);

require &dirname($0) . '\testsqllogin.pl';
require '..\helpers\assemblies.pl';

use vars qw($verbose $sql @result $type $prec $scale $len $no_of_tests
            $clr_enabled);

$verbose = shift @ARGV;

$^W = 1;

$| = 1;

my $X = testsqllogin();
my ($sqlver) = split(/\./, $X->{SQL_version});

print "1..31\n";

# Start with integer data types.
$sql = <<'SQLEND';
CREATE TABLE #a (a int NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
if ($type eq 'int') {
   print "ok 1\n";
}
else {
   print "not ok 1 # $type\n";
}

$sql = <<'SQLEND';
CREATE TABLE #a (a smallint NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
if ($type eq 'smallint') {
   print "ok 2\n";
}
else {
   print "not ok 2 # $type\n";
}

$sql = <<'SQLEND';
CREATE TABLE #a (a tinyint NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
if ($type eq 'tinyint') {
   print "ok 3\n";
}
else {
   print "not ok 3 # $type\n";
}

$sql = <<'SQLEND';
CREATE TABLE #a (a bit NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
if ($type eq 'bit') {
   print "ok 4\n";
}
else {
   print "not ok 4 # $type\n";
}

if ($sqlver >= 8) {
   $sql = <<'SQLEND';
   CREATE TABLE #a (a bigint NOT NULL)
   SELECT a FROM #a
   DROP TABLE #a
SQLEND
   @result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
   $type =  $result[0]->{'a'}{'Type'};
   if ($type eq 'bigint') {
      print "ok 5\n";
   }
   else {
      print "not ok 5 # $type\n";
   }
}
else {
   print "ok 5 # skip\n";
}

# Approxamite numeric types.
$sql = <<'SQLEND';
CREATE TABLE #a (a float NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
if ($type eq 'float') {
   print "ok 6\n";
}
else {
   print "not ok 6 # $type\n";
}

$sql = <<'SQLEND';
CREATE TABLE #a (a real NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
if ($type eq 'real') {
   print "ok 7\n";
}
else {
   print "not ok 7 # $type\n";
}

# Exact decimal types.
$sql = <<'SQLEND';
CREATE TABLE #a (a money NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
if ($type eq 'money') {
   print "ok 8\n";
}
else {
   print "not ok 8 # $type\n";
}

$sql = <<'SQLEND';
CREATE TABLE #a (a smallmoney NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
if ($type eq 'smallmoney') {
   print "ok 9\n";
}
else {
   print "not ok 9 # $type\n";
}

$sql = <<'SQLEND';
CREATE TABLE #a (a decimal(12, 7) NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
$prec =  $result[0]->{'a'}{'Precision'};
$scale = $result[0]->{'a'}{'Scale'};
if ($type eq 'decimal' and $prec == 12 and $scale == 7) {
   print "ok 10\n";
}
else {
   print "not ok 10 # <$type>  <$prec>  <$scale>\n";
}

# numeric. Note that this is reported as decimal.
$sql = <<'SQLEND';
CREATE TABLE #a (a numeric(23, 4) NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
$prec =  $result[0]->{'a'}{'Precision'};
$scale = $result[0]->{'a'}{'Scale'};
if (($type eq 'decimal' or $type eq 'numeric') and
     $prec == 23 and $scale == 4) {
   print "ok 11\n";
}
else {
   print "not ok 11 # <$type>  <$prec>  <$scale>\n";
}

# Datetime types
$sql = <<'SQLEND';
CREATE TABLE #a (a datetime NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
if ($type eq 'datetime') {
   print "ok 12\n";
}
else {
   print "not ok 12 # <$type>\n";
}

$sql = <<'SQLEND';
CREATE TABLE #a (a smalldatetime NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
if ($type eq 'smalldatetime') {
   print "ok 13\n";
}
else {
   print "not ok 13 # <$type>\n";
}

# binary data types
$sql = <<'SQLEND';
CREATE TABLE #a (a binary(19) NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
$len  =  $result[0]->{'a'}{'Maxlength'};
if ($type eq 'binary' and $len == 19) {
   print "ok 14\n";
}
else {
   print "not ok 14 # <$type> <$len>\n";
}

$sql = <<'SQLEND';
CREATE TABLE #a (a varbinary(125) NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
$len  =  $result[0]->{'a'}{'Maxlength'};
if ($type eq 'varbinary' and $len == 125) {
   print "ok 15\n";
}
else {
   print "not ok 15 # <$type> <$len>\n";
}

$sql = <<'SQLEND';
CREATE TABLE #a (a timestamp NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
if ($type eq 'timestamp') {
   print "ok 16\n";
}
else {
   print "not ok 16 # <$type> <$len>\n";
}

$sql = <<'SQLEND';
CREATE TABLE #a (a image NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
$len  =  $result[0]->{'a'}{'Maxlength'};
if ($type eq 'varbinary' and not defined $len) {
   print "ok 17\n";
}
else {
   print "not ok 17 # <$type> <$len>\n";
}

if ($sqlver >= 9) {
   $sql = <<'SQLEND';
   CREATE TABLE #a (a varbinary(MAX) NOT NULL)
   SELECT a FROM #a
   DROP TABLE #a
SQLEND
   @result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
   $type =  $result[0]->{'a'}{'Type'};
   $len  =  $result[0]->{'a'}{'Maxlength'};
   if ($type eq 'varbinary' and not defined $len) {
      print "ok 18\n";
   }
   else {
      print "not ok 18 # <$type> <$len>\n";
   }
}
else {
   print "ok 18 # skip\n";
}

# Character data types
$sql = <<'SQLEND';
CREATE TABLE #a (a char(23) NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
$len  =  $result[0]->{'a'}{'Maxlength'};
if ($type eq 'char' and $len == 23) {
   print "ok 19\n";
}
else {
   print "not ok 19 # <$type> <$len>\n";
}

$sql = <<'SQLEND';
CREATE TABLE #a (a varchar(53) NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
$len  =  $result[0]->{'a'}{'Maxlength'};
if ($type eq 'varchar' and $len == 53) {
   print "ok 20\n";
}
else {
   print "not ok 20 # <$type> <$len>\n";
}

$sql = <<'SQLEND';
CREATE TABLE #a (a text NOT NULL)
SELECT a FROM #a
DROP TABLE #a
SQLEND
@result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
$type =  $result[0]->{'a'}{'Type'};
$len  =  $result[0]->{'a'}{'Maxlength'};
if ($type eq 'varchar' and not defined $len) {
   print "ok 21\n";
}
else {
   print "not ok 21 # <$type> <$len>\n";
}

if ($sqlver >= 9) {
   $sql = <<'SQLEND';
   CREATE TABLE #a (a varchar(MAX) NOT NULL)
   SELECT a FROM #a
   DROP TABLE #a
SQLEND
   @result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
   $type =  $result[0]->{'a'}{'Type'};
   $len  =  $result[0]->{'a'}{'Maxlength'};
   if ($type eq 'varchar' and not defined $len) {
      print "ok 22\n";
   }
   else {
      print "not ok 22 # <$type> <$len>\n";
   }
}
else {
   print "ok 22 #skip\n";
}

if ($sqlver >= 7) {
   $sql = <<'SQLEND';
   CREATE TABLE #a (a nchar(1) NOT NULL)
   SELECT a FROM #a
   DROP TABLE #a
SQLEND
   @result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
   $type =  $result[0]->{'a'}{'Type'};
   $len  =  $result[0]->{'a'}{'Maxlength'};
   if ($type eq 'nchar' and $len == 1) {
      print "ok 23\n";
   }
   else {
      print "not ok 23 # <$type> <$len>\n";
   }

   $sql = <<'SQLEND';
   CREATE TABLE #a (a nvarchar(7) NOT NULL)
   SELECT a FROM #a
   DROP TABLE #a
SQLEND
   @result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
   $type =  $result[0]->{'a'}{'Type'};
   $len  =  $result[0]->{'a'}{'Maxlength'};
   if ($type eq 'nvarchar' and $len == 7) {
      print "ok 24\n";
   }
   else {
      print "not ok 24 # <$type> <$len>\n";
   }

   $sql = <<'SQLEND';
   CREATE TABLE #a (a ntext NOT NULL)
   SELECT a FROM #a
   DROP TABLE #a
SQLEND
   @result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
   $type =  $result[0]->{'a'}{'Type'};
   $len  =  $result[0]->{'a'}{'Maxlength'};
   if ($type eq 'nvarchar' and not defined $len) {
      print "ok 25\n";
   }
   else {
      print "not ok 25 # <$type> <$len>\n";
   }
}
else {
   print "ok 23 # skip\n";
   print "ok 24 # skip\n";
   print "ok 25 # skip\n";
}

if ($sqlver >= 9) {
   $sql = <<'SQLEND';
   CREATE TABLE #a (a nvarchar(MAX) NOT NULL)
   SELECT a FROM #a
   DROP TABLE #a
SQLEND
   @result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
   $type =  $result[0]->{'a'}{'Type'};
   $len  =  $result[0]->{'a'}{'Maxlength'};
   if ($type eq 'nvarchar' and not defined $len) {
      print "ok 26\n";
   }
   else {
      print "not ok 26 # <$type> <$len>\n";
   }
}
else {
   print "ok 26 # skip\n";
}

# GUID
if ($sqlver >= 7) {
   $sql = <<'SQLEND';
   CREATE TABLE #a (a uniqueidentifier NOT NULL)
   SELECT a FROM #a
   DROP TABLE #a
SQLEND
   @result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
   $type =  $result[0]->{'a'}{'Type'};
   if ($type eq 'uniqueidentifier') {
      print "ok 27\n";
   }
   else {
      print "not ok 27 # <$type>\n";
   }
}
else {
   print "ok 27 # skip\n";
}

# sql_variant.
if ($sqlver >= 8) {
   $sql = <<'SQLEND';
   CREATE TABLE #a (a sql_variant NOT NULL)
   SELECT a FROM #a
   DROP TABLE #a
SQLEND
   @result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
   $type =  $result[0]->{'a'}{'Type'};
   if ($type eq 'sql_variant') {
      print "ok 28\n";
   }
   else {
      print "not ok 28 # <$type>\n";
   }
}
else {
   print "ok 28 # skip\n";
}

# XML
if ($sqlver >= 9 and $X->{Provider} >= PROVIDER_SQLNCLI) {
   $sql = <<'SQLEND';
   CREATE TABLE #a (a xml NOT NULL)
   SELECT a FROM #a
   DROP TABLE #a
SQLEND
   @result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
   $type =  $result[0]->{'a'}{'Type'};
   if ($type eq 'xml') {
      print "ok 29\n";
   }
   else {
      print "not ok 29 # <$type>\n";
   }

   $X->sql(<<'SQLEND');
   IF NOT EXISTS (SELECT * FROM sys.xml_schema_collections WHERE name = 'OlleSC')
      CREATE XML SCHEMA COLLECTION OlleSC AS '
      <schema xmlns="http://www.w3.org/2001/XMLSchema">
            <element name="root" type="string"/>
      </schema>
      '
SQLEND
   $sql = <<'SQLEND';
   CREATE TABLE #a (a xml(OlleSC) NOT NULL)
   SELECT a FROM #a
   DROP TABLE #a
   DROP XML SCHEMA COLLECTION OlleSC
SQLEND
   @result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
   $type =  $result[0]->{'a'}{'Type'};
   if ($type eq 'xml') {
      print "ok 30\n";
   }
   else {
      print "not ok 30 # <$type> <$len>\n";
   }
}
else {
   print "ok 29 # skip\n";
   print "ok 30 # skip\n";
}

# CLR UDTs
if ($sqlver >= 9 and $X->{Provider} >= PROVIDER_SQLNCLI) {
   $clr_enabled = sql_one(<<SQLEND, Win32::SqlServer::SCALAR);
   SELECT value
   FROM   sys.configurations
   WHERE  name = 'clr enabled'
SQLEND
}

if ($clr_enabled) {
   create_the_udts($X, 'OlleComplexInteger', 'OllePoint', 'OlleString');
   $sql = <<'SQLEND';
   CREATE TABLE #a (a OlleComplexInteger NOT NULL)
   SELECT a FROM #a
   DROP TABLE #a
SQLEND
   @result = $X->sql($sql, Win32::SqlServer::COLINFO_FULL);
   $type =  $result[0]->{'a'}{'Type'};
   if ($type eq 'UDT') {
      print "ok 31\n";
   }
   else {
      print "not ok 31 # <$type>\n";
   }
   delete_the_udts($X);

}
else {
   print "ok 31 # skip\n";
}


exit;
