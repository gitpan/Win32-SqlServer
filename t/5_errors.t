#---------------------------------------------------------------------
# $Header: /Perl/OlleDB/t/5_errors.t 13    05-11-26 23:47 Sommar $
#
# Tests sql_message_handler and errors raised by OlleDB itself.
#
# $History: 5_errors.t $
# 
# *****************  Version 13  *****************
# User: Sommar       Date: 05-11-26   Time: 23:47
# Updated in $/Perl/OlleDB/t
# Renamed the module from MSSQL::OlleDB to Win32::SqlServer.
#
# *****************  Version 12  *****************
# User: Sommar       Date: 05-11-13   Time: 21:46
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 11  *****************
# User: Sommar       Date: 05-08-17   Time: 0:37
# Updated in $/Perl/OlleDB/t
# Keys in Messages entries is now in uppercase.
#
# *****************  Version 10  *****************
# User: Sommar       Date: 05-08-14   Time: 19:55
# Updated in $/Perl/OlleDB/t
# Added tests for SQLstate.
#
# *****************  Version 9  *****************
# User: Sommar       Date: 05-08-09   Time: 21:21
# Updated in $/Perl/OlleDB/t
# Msg_handler is now MsgHandler.
#
# *****************  Version 8  *****************
# User: Sommar       Date: 05-07-25   Time: 0:40
# Updated in $/Perl/OlleDB/t
# Added  tests for errors with UDT and XML.
#
# *****************  Version 7  *****************
# User: Sommar       Date: 05-06-27   Time: 21:41
# Updated in $/Perl/OlleDB/t
# Do prepend output file with directory name; testsqllogin.pl will chdir
# to the test directory.
#
# *****************  Version 6  *****************
# User: Sommar       Date: 05-06-26   Time: 22:36
# Updated in $/Perl/OlleDB/t
# Adapted to some changes in parameter handling.
#
# *****************  Version 5  *****************
# User: Sommar       Date: 05-03-28   Time: 18:42
# Updated in $/Perl/OlleDB/t
# Added test for too many parameters for a procedure.
#
# *****************  Version 4  *****************
# User: Sommar       Date: 05-02-27   Time: 22:54
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 3  *****************
# User: Sommar       Date: 05-02-27   Time: 21:54
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 2  *****************
# User: Sommar       Date: 05-02-27   Time: 17:44
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 1  *****************
# User: Sommar       Date: 05-02-20   Time: 23:12
# Created in $/Perl/OlleDB/t
#---------------------------------------------------------------------

use strict;
use Win32::SqlServer qw(:DEFAULT :consts);

use FileHandle;
use IO::Handle;
use File::Basename qw(dirname);

require &dirname($0) . '\testsqllogin.pl';

my($sql, $sql_call, $sp_call, $sql_callback, $msg_part, $sp_sql,
   $msgtext, $linestart, $expect_print, $expect_msgs, $errno, $state);

sub setup_a_test {
   # Sets up a test with an SQL command, an SP call and some stuff to be expected.
   my($sev) = @_;
   $errno     = 50000;
   $state     = 12;
   $msgtext   = "Er geht an die Ecke.";
   $sql       = qq!RAISERROR('$msgtext', $sev, $state)!;
   $sql_call  = "sql(q!$sql!, NORESULT)";
   $sp_call   = "sql_sp('#nisse_sp', ['$msgtext', $sev])";
   $msg_part  = "SQL Server message $errno, Severity $sev, State $state(, Server .+)?";
   $linestart = ' {3,5}1> ';
   $sp_sql    = "EXEC #nisse_sp (\\\@msgtext = '$msgtext', \\\@sev = $sev)";
}

$^W = 1;
$| = 1;

print "1..141\n";


my $X = testsqllogin();
my $sqlver = (split(/\./, $X->{SQL_version}))[0];
$X->{ErrInfo}{CheckRetStat} = 0;

$X->sql(<<'SQLEND');
   CREATE PROCEDURE #nisse_sp @msgtext varchar(25), @sev int AS
   RAISERROR(@msgtext, @sev, 12)
SQLEND

# Default setting for error. Should die and print it all.
setup_a_test(11);
$expect_print = ["=~ /^$msg_part\\n/i",
                 "=~ /Procedure\\s+#nisse_sp[_0-9A-F]+,\\s+Line 2/",
                 "eq '$msgtext\n'",
                 "=~ /$linestart$sp_sql\E\n/"];
do_test($sp_call, 1, 1, $expect_print);

# Default setting for warning message. Should print message details but not
# lines, and not die.
setup_a_test(9);
$expect_print = ["=~ /^$msg_part\\n/i",
                 "=~ /Procedure\\s+#nisse_sp[_0-9A-F]+,\\s+Line 2/",
                 "eq '$msgtext\n'"];
do_test($sp_call, 4, 0, $expect_print);

# Default setting for print message. Should print only message, and not die.
setup_a_test(0);
$expect_print = ["eq '$msgtext\n'"];
do_test($sp_call, 7, 0, $expect_print);

# This should be completely silent.
do_test("sql('USE master')", 10, 0, []);

# But this should print the message.
delete $X->{ErrInfo}{NeverPrint}{5701};
do_test("sql('USE tempdb')", 13, 0, ["=~ /Changed database context/i"]);

# Again an error, but should not print lines, and not abort. But there should
# be a Perl warning.
setup_a_test(11);
$X->{errInfo}{neverStopOn}{$errno}++;
$X->{errInfo}{printLines} = 12;
$expect_print = ["=~ /^$msg_part\n/i",
                 "eq 'Line 1\n'",
                 "eq '$msgtext\n'",
                 "=~ /Message from SQL Server at/"];
do_test($sql_call, 16, 0, $expect_print);

# Should print full text. Should not abort. Should return messages.
setup_a_test(7);
$X->{errInfo}{neverStopOn}{$errno} = 0;
$X->{errInfo}{maxSeverity} = 7;
$X->{errInfo}{alwaysPrint}{$errno}++;
$X->{errInfo}{saveMessages}++;
$expect_print = ["=~ /^$msg_part\n/i",
                 "eq 'Line 1\n'",
                 "eq '$msgtext\n'",
                 "=~ /$linestart\Q$sql\E\n/",
                 "=~ /Message from SQL Server at/"];
$expect_msgs = [{State    => "== $state",
                 Errno    => "== $errno",
                 Severity => "== 7",
                 Text     => "eq '$msgtext'",
                 Line     => "== 1",
                 Server   => 'or 1',
                 SQLstate => '=~ /^\w{5,5}$/'}];
do_test($sql_call, 19, 0, $expect_print, $expect_msgs);

# Should abort. Should not print. Should not return new messages, but keep old.
setup_a_test(9);
$X->{errInfo}{alwaysStopOn}{$errno}++;
$X->{errInfo}{alwaysPrint} = 0;
$X->{errInfo}{neverPrint}{$errno}++;
$X->{errInfo}{saveMessages} = 0;
do_test($sp_call, 22, 1, [], $expect_msgs);

# Should abort. Should only print the text. Should not return messages.
delete $X->{errInfo}{alwaysPrint};
delete $X->{errInfo}{neverPrint};
delete $X->{errInfo}{messages};
$X->{errInfo}{printMsg} = 10;
$X->{errInfo}{CarpLevel} = 9;
$expect_print = ["eq '$msgtext\n'"];
do_test($sp_call, 25, 1, $expect_print);

# Should not abort. Should print the text and a Perl warning.
$X->{errInfo}{MaxSeverity} = 11;
delete $X->{errInfo}{alwaysStopOn}{$errno};
$X->{errInfo}{CarpLevel} = 9;
$expect_print = ["eq '$msgtext\n'",
                 "=~ /Message from SQL Server at/"];
do_test($sp_call, 28, 0, $expect_print);

# We now test the default XS handler.
setup_a_test(11);
$X->{MsgHandler} = undef;
$expect_print = [q!=~ /^(Server .+, )?Msg 50000, Level 11, State 12, Procedure '#nisse_sp[_0-9a-fA-F]+', Line 2/!,
                 "=~ /\\s+$msgtext\\n/"];
do_test($sp_call, 31, 0, $expect_print);

# And for informational message.
setup_a_test(9);
$X->{MsgHandler} = undef;
$expect_print = ["=~ /^$msgtext/"];
do_test($sp_call, 34, 0, $expect_print);

# Now we test to use a customer msg handler.
sub custom_MsgHandler {
   my ($X, $errno, $state, $sev, $text) = @_;
   print STDERR "This is the message: '$text'\n";
   return ($sev <= 10);
}
$X->{MsgHandler} = \&custom_MsgHandler;
$expect_print = [qq!eq "This is the message: '$msgtext'\n"!];
do_test($sp_call, 37, 0, $expect_print);

# Now it should abort
setup_a_test(11);
do_test($sp_call, 40, 1, $expect_print);

# Restore defaults by dropping $X and recreate.
undef $X;
$X = testsqllogin();

# We will now test settings for SQL state. First default.
$sql_call = q!$X->sql("WAITFOR DELAY '00:00:05'", NORESULT)!;
$X->{CommandTimeout} = 1;
$expect_print =
    ['=~ /Message HYT00 .* (OLE DB Provider|Native Client)/',
     qq!=~ /[Tt]imeout expired\n/!,
     "=~ / 1> WAITFOR DELAY '00:00:05'/"];
do_test($sql_call, 43, 1, $expect_print);

# Suppress message.
$X->{ErrInfo}{NeverPrint}{'HYT00'}++;
do_test($sql_call, 46, 1, []);

# And don't die.
$X->{ErrInfo}{NeverStopOn}{'HYT00'}++;
do_test($sql_call, 49, 0, []);

# Remove this, but raise level for when to print and severity.
delete  $X->{ErrInfo}{NeverStopOn}{'HYT00'};
delete  $X->{ErrInfo}{NeverPrint}{'HYT00'};
$X->{ErrInfo}{MaxSeverity} = 17;
$X->{ErrInfo}{PrintLines} = 17;
$expect_print =
    ['=~ /Message HYT00 .* (OLE DB Provider|Native Client)/',
     qq!=~ /[Tt]imeout expired\n/!,
     "=~ /Message from Microsoft (OLE DB Provider for SQL Server|SQL Native Client) at/"];
do_test($sql_call, 52, 0, $expect_print);

# Test AlwaysStopOn and AlwaysPrint.
$X->{ErrInfo}{AlwaysPrint}{'HYT00'}++;
$X->{ErrInfo}{AlwaysStopOn}{'HYT00'}++;
$expect_print =
    ['=~ /Message HYT00 .* (OLE DB Provider|Native Client)/',
     qq!=~ /[Tt]imeout expired\n/!,
     "=~ / 1> WAITFOR DELAY '00:00:05'/"];
do_test($sql_call, 55, 1, $expect_print);

# Once more restore defaults by dropping $X and recreate.
undef $X;
$X = testsqllogin();

# Now we test that if there are multiple errors that we get them all.
$X->{ErrInfo}{SaveMessages} = 1;
$sql = <<SQLEND;
CREATE TABLE #abc(a int NOT NULL)
DECLARE \@x int
INSERT #abc(a) VALUES (\@x)
SQLEND
$sql_call  = "\$X->sql(q!$sql!, NORESULT)";
my $nulltext = ($sqlver > 6 ? 'Cannot insert.+NULL' : 'Attempt to insert.+NULL');
my $termintext = ($sqlver > 6 ? 'The statement has been terminated' : 'Command has been aborted');
$expect_print =
    ['=~ /SQL Server message 515, Severity 1[1-6], State \d+(, Server .+)?/',
     qq!eq "Line 3\n"!,
     "=~ /$nulltext/",
     '=~ / 1> CREATE TABLE #abc/',
     '=~ / 2> DECLARE \@x int/',
     '=~ / 3> INSERT #abc/',
     "=~ /$termintext/"];
$expect_msgs = [{State    => ">= 1",
                 Errno    => "== 515",
                 Severity => ">= 11",
                 Text     => "=~ /$nulltext/",
                 Line     => "== 3",
                 Server   => 'or 1',
                 SQLstate => '=~ /^\w{5,5}$/'},
                {State    => ">= 1",
                 Errno    => "== 3621",
                 Severity => "== 0",
                 Text     => "=~ /$termintext/",
                 Server   => 'or 1',
                 Line     => "== 3",
                 SQLstate => 'eq "01000"'}];
do_test($sql_call, 58, 1, $expect_print, $expect_msgs);

# And the same test for Perl warnings.
$X->{ErrInfo}{MaxSeverity} = 17;
$X->{ErrInfo}{SaveMessages} = 0;
delete $X->{ErrInfo}{Messages};
$sql_call = "\$X->sql(q!INSERT #abc(a) VALUES (NULL)!, NORESULT)";
$expect_print =
    ['=~ /SQL Server message 515, Severity 1[1-6], State \d+(, Server .+)?/',
     qq!eq "Line 1\n"!,
     '=~ /$nulltext/',
     '=~ / 1> INSERT #abc/',
     "=~ /$termintext/",
     "=~ /Message from SQL Server at/"];
do_test($sql_call, 61, 0, $expect_print);

# Next we will text the LinesWindow feature.
$X->{ErrInfo}{MaxSeverity} = 10;
$X->{ErrInfo}{SaveMessages} = 0;
delete $X->{ErrInfo}{Messages};
$sql = <<SQLEND;
-- 1st line.
-- 2nd line.
-- 3rd line.
-- 4th line.
RAISERROR('This is where it goes wrong', 11, 1)
-- 6th line.
-- 7th line.
SQLEND
$sql_call  = "\$X->sql(q!$sql!, NORESULT)";
$msg_part  = "SQL Server message 50000, Severity 11, State 1(, Server .+)?";
$expect_print = ["=~ /^$msg_part\\n/i",
                 "eq 'Line 5\n'",
                 "eq 'This is where it goes wrong\n'",
                 '=~ / 1> -- 1st line\.\n$/',
                 '=~ / 2> -- 2nd line\.\n$/',
                 '=~ / 3> -- 3rd line\.\n$/',
                 '=~ / 4> -- 4th line\.\n$/',
                q!=~ / 5> RAISERROR\('This is where it goes wrong', 11, 1\)\n$/!,
                 '=~ / 6> -- 6th line\.\n$/',
                 '=~ / 7> -- 7th line\.\n$/'];
do_test($sql_call, 64, 1, $expect_print);

$X->{ErrInfo}{LinesWindow} = 0;
$expect_print = ["=~ /^$msg_part\\n/i",
                 "eq 'Line 5\n'",
                 "eq 'This is where it goes wrong\n'",
         q!=~ / 5> RAISERROR\('This is where it goes wrong', 11, 1\)\n$/!];
do_test($sql_call, 67, 1, $expect_print);

$X->{ErrInfo}{LinesWindow} = 1;
$expect_print = ["=~ /^$msg_part\\n/i",
                 "eq 'Line 5\n'",
                 "eq 'This is where it goes wrong\n'",
                 '=~ /^ {3,5}4> -- 4th line\.\n$/',
         q!=~ / 5> RAISERROR\('This is where it goes wrong', 11, 1\)\n$/!,
                 '=~ /^ {3,5}6> -- 6th line\.\n$/'];
do_test($sql_call, 70, 1, $expect_print);

$X->{ErrInfo}{LinesWindow} = 3;
$expect_print = ["=~ /^$msg_part\\n/i",
                 "eq 'Line 5\n'",
                 "eq 'This is where it goes wrong\n'",
                 '=~ / 2> -- 2nd line\.\n$/',
                 '=~ / 3> -- 3rd line\.\n$/',
                 '=~ / 4> -- 4th line\.\n$/',
                q!=~ / 5> RAISERROR\('This is where it goes wrong', 11, 1\)\n$/!,
                 '=~ / 6> -- 6th line\.\n$/',
                 '=~ / 7> -- 7th line\.\n$/'];
do_test($sql_call, 73, 1, $expect_print);

# Now we test messages from the OLE DB provider. First one of these obscure
# message we can't really tell what they are due to.
$X->{ErrInfo}{SaveMessages} = 1;
$sql_call = <<'PERLEND';
$X->initbatch("SELECT ?");
$X->executebatch;
$X->cancelbatch;
PERLEND
$expect_print =
    ['=~ /^Message [0-9A-F]{8}.*(OLE DB Provider|Native Client).*Severity:? 16/i',
     '=~ /Win32::SqlServer call/',
     '=~ /No value given/',
     '=~ / 1> SELECT \?/'];
$expect_msgs = [{State    => "== 127",
                 Errno    => "<= -1",
                 Severity => "== 16",
                 Text     => "=~ /No value given/",
                 Line     => "== 0",
                 Proc     => "eq 'cmdtext_ptr->Execute'",
                 SQLstate => "=~ /[0-9A-F]{8}/i",
                 Source   => "=~ /OLE DB Provider|Native Client/"}];
do_test($sql_call, 76, 1, $expect_print, $expect_msgs);

# This one generates "Invalid character for cast specification"
$X->sql(<<'SQLEND');
   CREATE PROCEDURE #date_sp @d smalldatetime AS
   SELECT @d = @d
SQLEND
$sql_call = '$X->sql_sp("#date_sp", ["2103-01-01"])';
$expect_print =
    ['=~ /^Message \w{5}.*(OLE DB Provider|Native Client).*Severity:? 16/i',
     "=~ /Invalid.*cast specification/",
     q!=~ / {3,5}1> EXEC #date_sp\s+\@d\s*=\s*'2103-01-01'/!];
push(@$expect_msgs, {State    => '>= 1',
                     Errno    => '== 0',
                     Severity => '== 16',
                     Text     => '=~ /Invalid.*cast specification/',
                     Line     => '== 0',
                     SQLstate => '=~ /\w{5}/',
                     Source   => '=~ /OLE DB Provider|Native Client/'});
do_test($sql_call, 79, 1, $expect_print, $expect_msgs);

# Now we move on to test OlleDB's own messages. First errors with datetime
# hashes.
delete $X->{ErrInfo}{Messages};
$sql_call = '$X->sql_sp("#date_sp", [{Year => 1991, Day => 17}])';
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     "=~ /Mandatory part 'Month' missing/",
     "=~ /Message from Win32::SqlServer at/",
     '=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     "=~ /Could not convert Perl value.+smalldatetime/",
     "=~ /Message from Win32::SqlServer at/",
     '=~ /^Message -1.+Win32::SqlServer.+Severity:? 16/',
     "=~ /One or more parameters .+ Cannot execute/",
     q!=~ / 1> EXEC #date_sp\s+\@d\s*=\s*'HASH\(/!];
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => "=~ /Mandatory part 'Month' missing/",
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => "=~ /Could not convert Perl value.+smalldatetime/",
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 16',
                 Text     => "=~ /One or more parameters .+ Cannot execute/",
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"}];
do_test($sql_call, 82, 1, $expect_print, $expect_msgs);

# An error with an illegal value. We also test what happens if MaxSeverity
# permits continued execution.
delete $X->{ErrInfo}{Messages};
$X->{ErrInfo}{MaxSeverity} = 17;
$sql_call = '$X->sql_sp("#date_sp", [{Year => 1991, Month => 13, Day => 17}])';
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     "=~ /Part 'Month' .+ illegal value 13/",
     "=~ /Message from Win32::SqlServer at/",
     '=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     "=~ /Could not convert Perl value.+smalldatetime/",
     "=~ /Message from Win32::SqlServer at/",
     '=~ /^Message -1.+Win32::SqlServer.+Severity:? 16/',
     "=~ /One or more parameters .+ Cannot execute/",
     q!=~ / 1> EXEC #date_sp\s+\@d\s*=\s*'HASH\(/!,
     "=~ /Message from Win32::SqlServer at/"];
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => "=~ /Part 'Month' .+ illegal value 13/",
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => "=~ /Could not convert Perl value.+smalldatetime/",
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 16',
                 Text     => "=~ /One or more parameters .+ Cannot execute/",
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"}];
do_test($sql_call, 85, 0, $expect_print, $expect_msgs);

# Unknown data type and an illegal decimal value.
delete $X->{ErrInfo}{Messages};
$X->{ErrInfo}{MaxSeverity} = 10;
$sql_call = <<'PERLEND';
$X->sql('SELECT ?, ?', [['bludder', 12],
                        ['decimal(5,2)', 12345]]);
PERLEND
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     q!=~ /'bludder' .+ parameter '\@P1' .*illegal/!,
     "=~ /Message from Win32::SqlServer at/",
     '=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     "=~ /Could not convert Perl value .+12345.+ decimal/",
     "=~ /Message from Win32::SqlServer at/",
     '=~ /^Message -1.+Win32::SqlServer.+Severity:? 16/',
     "=~ /One or more parameters .+ Cannot execute/"];
if ($sqlver > 6) {
   push(@$expect_print,
         q!=~ / 1> EXEC sp_executesql\s+N'SELECT \@P1, \@P2'/!,
         q!=~ / 2> \s+N'\@P1 bludder,\s+\@P2 decimal\(5,\s*2\)',/!,
         q!=~ / 3> \s+\@P1\s*=\s*12,\s+\@P2\s*=\s*12345\s/!);
}
else {
   push(@$expect_print,
         q!=~ / 1> SELECT \?, \?/!,
         q!=~ / 2> \/\*\s+N'\@P1 bludder,\s+\@P2 decimal\(5,\s*2\)',/!,
         q!=~ / 3> \s+\@P1\s*=\s*12,\s+\@P2\s*=\s*12345\s*\*\//!);
}
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => q!=~ /'bludder' .+ parameter '\@P1' .*illegal/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => "=~ /Could not convert Perl value .+12345.+ decimal/",
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 16',
                 Text     => "=~ /One or more parameters .+ Cannot execute/",
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"}];
do_test($sql_call, 88, 1, $expect_print, $expect_msgs);

# Malformed data types
delete $X->{ErrInfo}{Messages};
$sql_call = <<'PERLEND';
$X->sql('SELECT ?, ?', [['float(53)', 12],
                        ['binary(5,2)', 12345]]);
PERLEND
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     q!=~ /'float\(53\)' .+ parameter '\@P1' .*illegal/!,
     "=~ /Message from Win32::SqlServer at/",
     '=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     q!=~ /'binary\(5,2\)' .+ parameter '\@P2' .*illegal/!,
     "=~ /Message from Win32::SqlServer at/",
     '=~ /^Message -1.+Win32::SqlServer.+Severity:? 16/',
     "=~ /One or more parameters .+ Cannot execute/"];
if ($sqlver > 6) {
   push(@$expect_print,
         q!=~ / 1> EXEC sp_executesql\s+N'SELECT \@P1, \@P2'/!,
         q!=~ / 2> \s+N'\@P1 float\(53\), \@P2 binary\(5,2\)',/!,
         q!=~ / 3> \s+\@P1 = 12, \@P2 = 12345\s/!);
}
else {
   push(@$expect_print,
         q!=~ / 1> SELECT \?, \?/!,
         q!=~ / 2> \/\*\s+N'\@P1 float\(53\), \@P2 binary\(5,2\)',/!,
         q!=~ / 3> \s+\@P1 = 12, \@P2 = 12345\s*\*\//!);
}
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => q!=~ /'float\(53\)' .+ parameter '\@P1' .*illegal/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => q!=~ /'binary\(5,2\)' .+ parameter '\@P2' .*illegal/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 16',
                 Text     => "=~ /One or more parameters .+ Cannot execute/",
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"}];
do_test($sql_call, 91, 1, $expect_print, $expect_msgs);


# Testing call to non-existing stored procedure
delete $X->{ErrInfo}{Messages};
$sql_call = q!$X->sql_sp('#notthere')!;
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 16/',
     "=~ /procedure '#notthere'/"];
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 16',
                 Text     => q!=~ /procedure '#notthere'/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"}];
do_test($sql_call, 94, 1, $expect_print, $expect_msgs);

# Test calling procedure with non-existing parameter.
delete $X->{ErrInfo}{Messages};
$sql_call = '$X->sql_sp("#date_sp", {notpar => "2103-01-01", hugo => 12})';
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     q!=~ /Procedure '#date_sp' .+ parameter.+'\@(notpar|hugo)'/!,
     "=~ /Message from Win32::SqlServer at/",
     q'=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     q!=~ /Procedure '#date_sp' .+ parameter.+'\@(notpar|hugo)'/!,
     "=~ /Message from Win32::SqlServer at/",
     q'=~ /^Message -1.+Win32::SqlServer.+Severity:? 16/',
     q!=~ /2 unknown.+Cannot execute/!];
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => q!=~ /Procedure '#date_sp' .+ parameter.+'\@(notpar|hugo)'/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => q!=~ /Procedure '#date_sp' .+ parameter.+'\@(notpar|hugo)'/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 16',
                 Text     => q!=~ /2 unknown.+Cannot execute/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"}];
do_test($sql_call, 97, 1, $expect_print, $expect_msgs);

# Test calling procedure with too many parameters.
delete $X->{ErrInfo}{Messages};
$sql_call = '$X->sql_sp("#date_sp", ["2103-01-01", 12])';
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 16/',
     q!=~ /2 parameters passed .+ '#date_sp' .+ only .*(one|1) parameter\b/!];
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 16',
                 Text     => q!=~ /2 parameters passed .+ '#date_sp' .+ only .*(one|1) parameter\b/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"}];
do_test($sql_call, 100, 1, $expect_print, $expect_msgs);


# Named/unnamed mixup, and OUTPUT is not reference. This includes a PRINT
# messages to see that the correct value of @sev is used.
$X->sql(<<'SQLEND');
   CREATE PROCEDURE #partest_sp @msgtext varchar(25) OUTPUT,
                                @sev int, @r int = 9 OUTPUT AS
   RAISERROR(@msgtext, @sev, 12)
SQLEND
delete $X->{ErrInfo}{Messages};
$sql_call = '$X->sql_sp("#partest_sp", ["Plain vanilla", 0, \$state], {sev => 17})';
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     q!=~ /arameter '\@sev' .+ position 2 .+ unnamed .+ named/!,
     "=~ /Message from Win32::SqlServer at/",
     '=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     q!=~ /Output parameter '\@msgtext' .+ not .+ reference/!,
     "=~ /Message from Win32::SqlServer at/",
     'eq "Plain vanilla\n"'];
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => q!=~ /arameter '\@sev' .+ position 2 .+ unnamed .+ named/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => q!=~ /Output parameter '\@msgtext' .+ not .+ reference/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '== 12',
                 Errno    => '== 50000',
                 Severity => '== 0',
                 Text     => 'eq "Plain vanilla"',
                 SQLstate => 'eq "01000"',
                 Server   => 'or 1',
                 Line     => '== 3',
                 Proc     => '=~ /^#partest_sp[_[0-9A-F]+/'}];
do_test($sql_call, 103, 0, $expect_print, $expect_msgs);


# Same parameter with and without the @. And test NoWhine. No warning about
# @msgtext here.
delete $X->{ErrInfo}{Messages};
$X->{ErrInfo}{NoWhine}++;
$sql_call = q!$X->sql_sp("#partest_sp", ["Plain vanilla"], {sev => 17, '@sev' => 0})!;
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     q!=~ /hash parameters .+ key 'sev' .+ '\@sev'/!,
     "=~ /Message from Win32::SqlServer at/",
     'eq "Plain vanilla\n"'];
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => q!=~ /hash parameters .+ 'sev' .+ '\@sev'/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '== 12',
                 Errno    => '== 50000',
                 Severity => '== 0',
                 Text     => 'eq "Plain vanilla"',
                 SQLstate => 'eq "01000"',
                 Server   => 'or 1',
                 Line     => '== 3',
                 Proc     => '=~ /^#partest_sp[_[0-9A-F]+/'}];
do_test($sql_call, 106, 0, $expect_print, $expect_msgs);


# sql_insert with non-existing table.
delete $X->{ErrInfo}{Messages};
$sql_call = '$X->sql_insert("#notthere", {sev => 17})';
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 16/',
     "=~ /Table '#notthere'/"];
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 16',
                 Text     => q!=~ /Table '#notthere'/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"}];
do_test($sql_call, 109, 1, $expect_print, $expect_msgs);

# Parameter clash with sql. Not tested on 6.5, since named parameters are not
# supported there.
delete $X->{ErrInfo}{Messages};
$sql_call = <<'PERLEND';
   $X->sql('RAISERROR(?, ?, ?)', [['varchar', 'This is jazz'],
                                  ['smallint', 14],
                                  ['smallint', 17]],
                                 {'@P2' => ['smallint', 10]});
PERLEND
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     q!=~ /named parameter '\@P2', .+ 3.+unnamed/!,
     "=~ /Message from Win32::SqlServer at/",
     '=~ /SQL Server message 50000, Severity 14, State 17(, Server .+)?/',
     qq!eq "Line 1\n"!,
     'eq "This is jazz\n"',
     q!=~ / 1> EXEC sp_executesql\s+N'RAISERROR\(\@P1, \@P2, \@P3\)'/!,
     q!=~ / 2> \s+N'\@P1 varchar\(\d+\),\s+\@P2 smallint,\s*\@P3 smallint',/!,
     q!=~ / 3> \s+\@P1 = 'This is jazz', \@P2 = 14, \@P3 = 17/!];
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => q!=~ /named parameter '\@P2', .+ 3.+unnamed/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '== 17',
                 Errno    => '== 50000',
                 Severity => '== 14',
                 Text     => 'eq "This is jazz"',
                 Line     => '== 1',
                 Server   => 'or 1',
                 SQLstate => 'eq "42000"'}];
do_test($sql_call, 112, 1, $expect_print, $expect_msgs, 1);

# No data type specified. And an OLE DB error on 6.5.
delete $X->{ErrInfo}{Messages};
$sql_call = <<'PERLEND';
   $X->sql('RAISERROR(?, 12, 17)', ['This is jazz', undef]);
PERLEND
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     q!=~ /no datatype .+ parameter '\@P1', value 'This is jazz'/!,
     "=~ /Message from Win32::SqlServer at/"];
if ($sqlver > 6) {
   push(@$expect_print,
     '=~ /SQL Server message 50000, Severity 12, State 17(, Server .+)?/',
     qq!eq "Line 1\n"!,
     'eq "This is jazz\n"',
     q!=~ / 1> EXEC sp_executesql\s+N'RAISERROR\(\@P1, 12, 17\)'/!,
     q!=~ / 2> \s+N'\@P1 char\(\d+\),\s+\@P2 char(\(1\))?'/!,
     q!=~ / 3> \s+\@P1 = 'This is jazz', \@P2 = NULL/!);
}
else {
   push(@$expect_print,
     '=~ /^Message [0-9A-F]{8}.*(OLE DB Provider|Native Client).*Severity:? 16/i',
     '=~ /Win32::SqlServer call/',
     '=~ /Multiple-step/',
     '=~  / 1> RAISERROR\(\?, 12, 17\)/',
     q!=~ / 2> \/\*\s+N'\@P1 char\(\d+\),\s+\@P2 char(\(1\))?'/!,
     q!=~ / 3> \s+\@P1 = 'This is jazz', \@P2 = NULL\s*\*\//!);
}
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => q!=~ /no datatype .+ parameter '\@P1', value 'This is jazz'/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '== 17',
                 Errno    => '== 50000',
                 Severity => '== 12',
                 Text     => 'eq "This is jazz"',
                 Line     => '== 1',
                 Server   => 'or 1',
                 SQLstate => 'eq "42000"'}];
if ($sqlver == 6) {
   $$expect_msgs[1] = {State    => "== 127",
                       Errno    => "<= -1",
                       Severity => "== 16",
                       Text     => "=~ /Multiple-step/",
                       Line     => "== 0",
                       Proc     => "eq 'cmdtext_ptr->Execute'",
                       SQLstate => "=~ /[0-9A-F]{8}/i",
                       Source   => "=~ /OLE DB Provider|Native Client/"};
}
do_test($sql_call, 115, 1, $expect_print, $expect_msgs);

# Scale/precision missing for decimal.
delete $X->{ErrInfo}{Messages};
$sql = <<'SQLEND';
DECLARE @out varchar(200)
SELECT @out = convert(varchar, ?) + ' -- ' + convert(varchar, ?) + ' -- ' +
              convert(varchar, ?)
PRINT @out
SQLEND
$sql_call = <<PERLEND;
   \$X->sql(q!$sql!, [['decimal', 47.11],
                     ['numeric(9)', 47.11],
                     ['decimal(9,3)', 47.11]]);
PERLEND
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     q!=~ /Precision .+ scale missing .+ '\@P1'/!,
     "=~ /Message from Win32::SqlServer at/",
     '=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     q!=~ /Precision .+ scale missing .+ '\@P2'/!,
     "=~ /Message from Win32::SqlServer at/",
     'eq "47 -- 47 -- 47.110\n"'];
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => q!=~ /Precision .+ scale missing .+ '\@P1'/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => q!=~ /Precision .+ scale missing .+ '\@P2'/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '>= 1',
                 Errno    => '== 0',
                 Severity => '== 0',
                 Text     => 'eq "47 -- 47 -- 47.110"',
                 Line     => '== 4',
                 Server   => 'or 1',
                 SQLstate => 'eq "01000"'}];
do_test($sql_call, 118, 0, $expect_print, $expect_msgs);

# Samma parameternamn två gånger:
delete $X->{ErrInfo}{Messages};
$sql_call = <<'PERLEND';
   $X->sql('RAISERROR(@P1, 4, 1)', {P1    => ['varchar', 'This is jazz'],
                                    '@P1' => ['varchar', 'Katzenjammer']});
PERLEND
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 10/',
     q!=~ /hash parameters .+ key 'P1' .+ '\@P1'/!,
     "=~ /Message from Win32::SqlServer at/",
     '=~ /SQL Server message 50000, Severity 4, State 1(, Server .+)?/',
     qq!eq "Line 1\n"!,
     'eq "Katzenjammer\n"'];
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 10',
                 Text     => q!=~ /hash parameters .+ key 'P1' .+ '\@P1'/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"},
                {State    => '== 1',
                 Errno    => '== 50000',
                 Severity => '== 4',
                 Text     => 'eq "Katzenjammer"',
                 Line     => '== 1',
                 Server   => 'or 1',
                 SQLstate => 'eq "01000"'}];
do_test($sql_call, 121, 0, $expect_print, $expect_msgs, 1);

# Using UDT without specifying user-type. (We can to this on all platforms,
# because this is trapped early by OlleDB itself.)
delete $X->{ErrInfo}{Messages};
$sql_call = q!$X->sql('SELECT ?', [['UDT', undef]])!;
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 16/',
    q!=~ /No actual user type .+ UDT .+ '\@P1'/!];
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 16',
                 Text     => q!=~ /No actual user type .+ UDT .+ '\@P1'/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"}];
do_test($sql_call, 124, 1, $expect_print, $expect_msgs, 0);

# UDT with conflicting specifiers.
delete $X->{ErrInfo}{Messages};
$sql_call = q!$X->sql('SELECT ?', [['UDT(OllePoint)', undef, 'OlleString']])!;
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 16/',
    q!=~ /Conflicting .+ \('OllePoint' and 'OlleString'\) .+ '\@P1' .+ UDT/!];
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 16',
                 Text     => q!=~ /Conflicting .+ \('OllePoint' and 'OlleString'\) .+ '\@P1' .+ UDT/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"}];
do_test($sql_call, 127, 1, $expect_print, $expect_msgs, 0);

# XML with conflicting specifiers.
delete $X->{ErrInfo}{Messages};
$sql_call = q!$X->sql('SELECT ?', [['xml(OlleSC)', undef, 'OlleSC2']])!;
$expect_print =
    ['=~ /^Message -1.+Win32::SqlServer.+Severity:? 16/',
    q!=~ /Conflicting .+ \('OlleSC' and 'OlleSC2'\) .+ '\@P1' .+ xml/!];
$expect_msgs = [{State    => '>= 1',
                 Errno    => '<= -1',
                 Severity => '== 16',
                 Text     => q!=~ /Conflicting .+ \('OlleSC' and 'OlleSC2'\) .+ '\@P1' .+ xml/!,
                 Line     => "== 0",
                 Source   => "eq 'Win32::SqlServer'"}];
do_test($sql_call, 130, 1, $expect_print, $expect_msgs, 0);


# We will now test sql_has_errors. First get a default connection.
undef $X;
$X = testsqllogin();
$X->{ErrInfo}{MaxSeverity} = 17;
$X->{ErrInfo}{NeverPrint}{50000}++;
$X->sql("RAISERROR('Test', 11, 10)", NORESULT);
{
   my (@warns);
   local $SIG{__WARN__} = sub {push(@warns, $_[0])};

   # Since SaveMessages is off we should not get false back
   if (not $X->sql_has_errors) {
      print "ok 133\n";
   }
   else {
      print "not ok 133\n";
   }

   # ...but we should be warned.
   if (@warns) {
      print "ok 134\n";
   }
   else {
      print "not ok 134\n";
   }
}

$X->{ErrInfo}{SaveMessages} = 1;
$X->sql("RAISERROR('Test', 11, 10)", NORESULT);
if ($X->sql_has_errors) {
   print "ok 135\n";
}
else {
   print "not ok 135\n";
}


# The error should still be there.
if (exists $X->{ErrInfo}{Messages}) {
   print "ok 136\n";
}
else {
   print "not ok 136\n";
}

delete $X->{ErrInfo}{Messages};
$X->sql("RAISERROR('Test', 9, 10)", NORESULT);
if (not $X->sql_has_errors(1)) {
   print "ok 137\n";
}
else {
   print "not ok 137\n";
}

# The message should still be there, as we said to sql_has_errors.
if (scalar(@{$X->{ErrInfo}{Messages}}) == 1) {
   print "ok 138\n";
}
else {
   print "not ok 138\n";
}

# But after this.
$X->sql_has_errors();
if (not exists $X->{ErrInfo}{Messages}) {
   print "ok 139\n";
}
else {
   print "not ok 139\n";
}

$X->sql(<<SQLEND, NORESULT);
RAISERROR('Test1', 9, 10)
RAISERROR('Test2', 11, 10)
RAISERROR('Test3', 0, 10)
SQLEND
if ($X->sql_has_errors()) {
   print "ok 140\n";
}
else {
   print "not ok 140\n";
}

if (scalar(@{$X->{ErrInfo}{Messages}}) == 3) {
   print "ok 141\n";
}
else {
   print "not ok 141\n";
}



# That's enough!
exit;


sub do_test{
   my($test, $test_no, $expect_die, $expect_print, $expect_msgs, $skip65) = @_;

   my($savestderr, $errfile, $fh, $evalout, @carpmsgs);

   if ($sqlver == 6 and $skip65) {
      print "ok " . $test_no++ . " # skip\n";
      print "ok " . $test_no++ . " # skip\n";
      print "ok " . $test_no++ . " # skip\n";
      return;
   }

   # Get file name.
   $errfile = "error.$test_no";

   # To start with, we alter between writing to STDERR and using a ErrFileHandle.
   # Later we give up on ErrFileHandle, to know where the Perl warnings will
   # appear.
   if ($test_no % 2 == 0 or $test_no > 21) {
      delete $X->{errInfo}{errFileHandle};

      # Save STDERR so we can reopen.
      $savestderr = FileHandle->new_from_fd(*main::STDERR, "w") or die "Can't dup STDERR: $!\n";

      # Redirect STDERR to a file.
      open(STDERR, ">$errfile") or die "Can't redriect STDERR to '$errfile': $!\n";
      STDERR->autoflush;

      # Run the test. Must eval, it may die.
      eval($test);
      $evalout = $@;

      # Put STDERR back to were it was.
      open(STDERR, ">&" . $savestderr->fileno) or (print "Can't reopen STDERR: $!\n" and die);
      STDERR->autoflush;
   }
   else {
      # Test errFileHandle
      $fh = new FileHandle;
      $fh->open($errfile, "w") or die "Can't write to '$errfile': $!\n";
      $X->{errInfo}{errFileHandle} = $fh;

      # Must set up a handler to catch warnings.
      local $SIG{__WARN__} = sub{push(@carpmsgs, $_[0])};

      # Run the test. Must eval, it may die.
      eval($test);
      $evalout = $@;

      $fh->close;
   }

   # Now, read the error file.
   $fh = new FileHandle;
   $fh->open($errfile, "r") or die "Cannot read $errfile: $!\n";
   my @errfile = <$fh>;
   $fh->close;

   # Add the warnings to the error file (they are already there if we did use
   # the ErrFileHandle.
   push(@errfile, @carpmsgs);

   # Did the execution terminate by croak? And should it have?
   if ($expect_die) {
      if ($evalout and $evalout =~ /^Terminating.*fatal/i) {
         print "ok $test_no\n"
      }
      else {
         print "# evalout = '$evalout'\n" if defined $evalout;
         print "not ok $test_no\n";
      }
   }
   else {
      if (not $evalout) {
         print "ok $test_no\n"
      }
      else {
         print "# evalout = '$evalout'\n";
         print "not ok $test_no\n";
      }
   }
   $test_no++;

   # Compare output.
   if (compare(\@errfile, $expect_print)) {
      print "ok $test_no\n"
   }
   else {
      print "not ok $test_no\n";
   }
   $test_no++;

   # Then the messages.
   if (compare($X->{errInfo}{'messages'}, $expect_msgs)) {
      print "ok $test_no\n"
   }
   else {
      print "not ok $test_no\n";
   }
   $test_no++;
}



sub compare {
   my ($x, $y) = @_;

   my ($refx, $refy, $ix, $key, $result);

   $refx = ref $x;
   $refy = ref $y;

   if (not $refx and not $refy) {
      if (defined $x and defined $y) {
         $result = eval("q!$x! $y");
         warn "no match: <$x> <$y>" if not $result;
         return $result;
      }
      else {
         $result = (not defined $x and not defined $y);
         warn  'Left is ' . (defined $x ? "'$x'" : 'undefined') .
               ' and right is ' . (defined $y ? "'$y'" : 'undefined')
               if not $result;
         return $result
      }
   }
   elsif ($refx ne $refy) {
      return 0;
   }
   elsif ($refx eq "ARRAY") {
      if ($#$x != $#$y) {
         warn  "Left has upper index $#$x and right has upper index $#$y.";
         return 0;
      }
      elsif ($#$x >= 0) {
         foreach $ix (0..$#$x) {
            $result = compare($$x[$ix], $$y[$ix]);
            last if not $result;
         }
         return $result;
      }
      else {
         return 1;
      }
   }
   elsif ($refx eq "HASH") {
      my $nokeys_x = scalar(keys %$x);
      my $nokeys_y = scalar(keys %$y);
      if ($nokeys_x == $nokeys_y and $nokeys_x == 0) {
         return 1;
      }
      if ($nokeys_x > 0) {
         foreach $key (keys %$x) {
            if (not exists $$y{$key} and defined $$x{$key}) {
                warn "Left has key '$key' which is missing from right.";
                return 0;
            }
            $result = compare($$x{$key}, $$y{$key});
            last if not $result;
         }
      }
      return 0 if not $result;
      foreach $key (keys %$y) {
         if (not exists $$x{$key} and defined $$y{$key}) {
             warn "Right has key '$key' which is missing from left.";
             return 0;
         }
      }
      return $result;
   }
   elsif ($refx eq "SCALAR") {
      return compare($$x, $$y);
   }
   else {
      $result = ($x eq $y);
      warn "no match: <$x> <$y>" if not $result;
      return $result;
   }
}
