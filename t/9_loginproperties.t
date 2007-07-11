#---------------------------------------------------------------------
# $Header: /Perl/OlleDB/t/9_loginproperties.t 16    07-07-07 16:43 Sommar $
#
# This test suite tests that setloginproperty, Autoclose and CommandTimeout.
#
# $History: 9_loginproperties.t $
# 
# *****************  Version 16  *****************
# User: Sommar       Date: 07-07-07   Time: 16:43
# Updated in $/Perl/OlleDB/t
# Added support for specifying different providers.
#
# *****************  Version 15  *****************
# User: Sommar       Date: 07-06-10   Time: 21:45
# Updated in $/Perl/OlleDB/t
# When testing that pooling is off, permit errors for monitor, since the
# error level is 16 on Katmai.
#
# *****************  Version 14  *****************
# User: Sommar       Date: 05-11-26   Time: 23:47
# Updated in $/Perl/OlleDB/t
# Renamed the module from MSSQL::OlleDB to Win32::SqlServer.
#
# *****************  Version 13  *****************
# User: Sommar       Date: 05-11-06   Time: 20:48
# Updated in $/Perl/OlleDB/t
# Move CommantTimeout tests to A_rowsetprops.t
#
# *****************  Version 12  *****************
# User: Sommar       Date: 05-10-16   Time: 23:34
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 11  *****************
# User: Sommar       Date: 05-08-20   Time: 22:50
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 10  *****************
# User: Sommar       Date: 05-08-14   Time: 19:55
# Updated in $/Perl/OlleDB/t
# Added tests for DisconnectOn
#
# *****************  Version 9  *****************
# User: Sommar       Date: 05-08-11   Time: 22:52
# Updated in $/Perl/OlleDB/t
# Added tests for is_connected().
#
# *****************  Version 8  *****************
# User: Sommar       Date: 05-07-25   Time: 0:41
# Updated in $/Perl/OlleDB/t
# Reworked the test for Network Address.
#
# *****************  Version 7  *****************
# User: Sommar       Date: 05-06-27   Time: 22:36
# Updated in $/Perl/OlleDB/t
# Test for integrated security did not cater for the case connection was
# trusted, but user was not granted access.
#
# *****************  Version 6  *****************
# User: Sommar       Date: 05-06-27   Time: 21:40
# Updated in $/Perl/OlleDB/t
# Change directory to the test directory.
#
# *****************  Version 5  *****************
# User: Sommar       Date: 05-06-25   Time: 17:10
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 4  *****************
# User: Sommar       Date: 05-06-20   Time: 23:00
# Updated in $/Perl/OlleDB/t
# Added test for OldPassword.
#
# *****************  Version 3  *****************
# User: Sommar       Date: 05-05-29   Time: 22:24
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 2  *****************
# User: Sommar       Date: 05-05-29   Time: 21:30
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 1  *****************
# User: Sommar       Date: 05-05-23   Time: 0:36
# Created in $/Perl/OlleDB/t
#---------------------------------------------------------------------

use strict;
use Win32::SqlServer qw(:DEFAULT :consts);
use File::Basename qw(dirname);

# This test script reads OLLEDBTEST directly, because it uses more fields
# from it.
my ($olledbtest) = $ENV{'OLLEDBTEST'};
my ($mainserver, $mainuser, $mainpw,
    $secondserver, $seconduser, $secondpw, $provider);
($mainserver, $mainuser, $mainpw, $secondserver, $seconduser, $secondpw, $provider) =
     split(/;/, $olledbtest) if defined $olledbtest;

sub setup_testc {
    # Creates a test connection object with some common initial settings.
    my ($userpw) = @_;
    $userpw = 1 if not defined $userpw;
    my $testc = new Win32::SqlServer;
    $testc->{Provider} = $provider if defined $provider;
    $testc->setloginproperty('Server', $mainserver) if $mainserver;
    if ($userpw and $mainuser) {
       $testc->setloginproperty('Username', $mainuser);
    }
    if ($userpw and $mainpw) {
       $testc->setloginproperty('Password', $mainpw);
    }
    $testc->{ErrInfo}{PrintMsg}    = 17;
    $testc->{ErrInfo}{PrintLines}  = 17;
    $testc->{ErrInfo}{PrintText}   = 17;
    $testc->{ErrInfo}{MaxSeverity} = 17;
    $testc->{ErrInfo}{CarpLevel}   = 17;
    $testc->{ErrInfo}{SaveMessages} = 1;
    return $testc;
}

$^W = 1;

$| = 1;

chdir dirname($0);

print "1..33\n";

# Set up a monitor connection
my $monitor = sql_init($mainserver, $mainuser, $mainpw, undef, $provider);
my ($monitorsqlver) = split(/\./, $monitor->{SQL_version});

# This is the connection we use for tests.
my $testc;

# Reappears every now and then.
my $errmsg;

# First tests, do we handle various values for Integrated Security corectly?
# Default we should connect with Integrated security.
$testc = setup_testc(0);
$testc->connect();
if (not $testc->{ErrInfo}{Messages} or
    $testc->{ErrInfo}{Messages}[0]{'text'} =~
       /(trusted .+ connection)|(Login failed for .+\\.+)/) {
    print "ok 1\n";
}
else {
    print "not ok 1\n";
}
$testc->disconnect;

# Explicit enable with numeric value.
$testc = setup_testc(0);
$testc->setloginproperty('IntegratedSecurity', 1);
$testc->connect();
if (not $testc->{ErrInfo}{Messages} or
    $testc->{ErrInfo}{Messages}[0]{'text'} =~
       /(trusted .+ connection)|(Login failed for .+\\.+)/) {
    print "ok 2\n";
}
else {
    print "not ok 2\n";
}
$testc->disconnect;


# Test login with Integrated Security off and nothing else on.
$testc->setloginproperty('IntegratedSecurity', 0);
delete $testc->{ErrInfo}{Messages};
$testc->connect();
$errmsg = $testc->{ErrInfo}{Messages}[0]{'text'};
if ($errmsg =~ /Invalid authorization specification/)  {
   print "ok 3\n";
}
else {
   print "not ok 3 # $errmsg\n";
}
$testc->disconnect;

# Tests for Windows authentication and SQL authentication.
my %loginconfig = $monitor->sql_one("EXEC master..xp_loginconfig 'login mode'");

# Get a new test connection with a fresh start.
$testc = setup_testc(0);
$testc->setloginproperty('Username', 'Slabanketti');
$testc->setloginproperty('Password', 'Once in a lifetime');
$testc->connect();
$errmsg = $testc->{ErrInfo}{Messages}[0]{'text'};
if ($loginconfig{'config_value'} !~ /^Windows/) {
   if ($errmsg =~ /[Ll]ogin failed/) {
      print "ok 4\n";
   }
   else {
      print "not ok 4 # $errmsg\n";
   }
}
else {
   if ($errmsg  =~ /trusted .+ connection/) {
       print "ok 4\n";
   }
   else {
       print "not ok 4\n";
   }
}

# Test database property. And try autoconnect, while we're at it.
$testc = setup_testc;
$testc->{AutoConnect} = 1;
my $db = $testc->sql_one('SELECT db_name()');
# Do we have tempdb as default?
if ($db eq 'tempdb') {
   print "ok 5\n";
}
else {
   print "not ok 5 # $db\n";
}

# Test explicit database.
$testc->setloginproperty('Database', 'master');
$db = $testc->sql_one('SELECT db_name()');
if ($db eq 'master') {
   print "ok 6\n";
}
else {
   print "not ok 6 # $db\n";
}

# Test server property. Here we can't test the default value, as there might
# not be a local server. But we want to test what happens when we change
# servers.
if ($secondserver and $secondserver ne $mainserver) {
   $testc = setup_testc;
   $testc->{AutoConnect} = 1;

   # Set up connection to second server.
   $testc->setloginproperty('Server', $secondserver);
   if ($seconduser) {
      $testc->setloginproperty('Username', $seconduser);
      $testc->setloginproperty('Password', $secondpw);
   }
   else {
      $testc->setloginproperty('IntegratedSecurity', "SSPI");
   }

   # Get SQL version first thing we do.
   my $newsqlver = $testc->{SQL_version};
   my %thissqlver = $testc->sql_one("EXEC master..xp_msver 'Productversion'");
   if ($thissqlver{'Character_Value'} eq $newsqlver) {
      print "ok 7\n";
   }
   else {
      print "not ok 7\n";
   }

   # But did we really change servers? (We can't test for SERVERNAME, as
   # $secondserver may be an IP-address or an alias.)
   my $servername1 = $monitor->sql_one('SELECT @@servername', SCALAR);
   my $servername2 = $testc->sql_one('SELECT @@servername', SCALAR);
   if ($servername1 ne $servername2) {
      print "ok 8\n";
   }
   else {
      print "not ok 8\n";
   }

   # And change back. Now we execute a command before we look at SQL_version.
   $testc->setloginproperty('Server', $mainserver);
   if ($mainuser) {
      $testc->setloginproperty('Username', $mainuser);
      $testc->setloginproperty('Password', $mainpw);
   }
   else {
      $testc->setloginproperty('IntegratedSecurity', "SSPI");
   }
   %thissqlver = $testc->sql_one("EXEC master..xp_msver 'Productversion'");
   if ($thissqlver{'Character_Value'} eq $testc->{SQL_version}) {
      print "ok 9\n";
   }
   else {
      print "not ok 9\n";
   }
}
else {
   print "ok 7 # skip, no second server.\n";
   print "ok 8 # skip, no second server.\n";
   print "ok 9 # skip, no second server.\n";
}

# Time for a new object. We're going to test connection pooling now.
# Pooling should be on by default.
$testc = setup_testc;
$testc->connect;
$monitor->{ErrInfo}{PrintText} = 2;    # Suppress DBCC messages on 6.5
my $spid = $testc->sql_one('SELECT @@spid', SCALAR);
$testc->sql("PRINT 'Cornershot'");
$testc->disconnect;
my $inputbuffer;
$inputbuffer = $monitor->sql("DBCC INPUTBUFFER($spid) WITH NO_INFOMSGS",
                             SINGLESET, HASH);
my $colname = ($monitorsqlver == 6 ? 'Input Buffer' : 'EventInfo');
if ($$inputbuffer[0] and $$inputbuffer[0]{$colname} =~ /^PRINT 'Cornershot'/) {
   print "ok 10\n";
}
else {
   print "not ok 10 # $$inputbuffer[0]{$colname}";
}

# Pooling off.
$testc->setloginproperty('Pooling', 0);
$testc->connect;
$spid = $testc->sql_one('SELECT @@spid', SCALAR);
$testc->sql("PRINT 'Penalty kick'");
$testc->disconnect;
sleep(1); # Permit for the test connection to actually go away.
$monitor->{ErrInfo}{PrintMsg} = 2;    # Suppress DBCC messages on 6.5
$monitor->{ErrInfo}{SaveMessages} = 1;
$monitor->{ErrInfo}{MaxSeverity} = 16;
$monitor->{ErrInfo}{PrintMsg} = 17;
$monitor->{ErrInfo}{PrintLines} = 17;
$monitor->{ErrInfo}{PrintText} = 17;
$monitor->{ErrInfo}{CarpLevel} = 17;
delete $monitor->{ErrInfo}{Messages};
$inputbuffer = $monitor->sql("DBCC INPUTBUFFER($spid) WITH NO_INFOMSGS",
                             SINGLEROW, HASH);
if ($monitor->{ErrInfo}{Messages} and
    $monitor->{ErrInfo}{Messages}[0]{'text'} =~ /(Invalid SPID|does not process input)/i) {
   print "ok 11\n";
}
else {
   print "not ok 11\n";
}
$monitor->{ErrInfo}{MaxSeverity} = 10;
$monitor->{ErrInfo}{PrintMsg} = 1;
$monitor->{ErrInfo}{PrintLines} = 11;
$monitor->{ErrInfo}{PrintText} = 0;
$monitor->{ErrInfo}{CarpLevel} = 11;


# Pooling on again
$testc->setloginproperty('Pooling', 1);
$testc->connect;
$spid = $testc->sql_one('SELECT @@spid', SCALAR);
$testc->sql("PRINT 'Elfmeter'");
$testc->disconnect;
$monitor->{ErrInfo}{PrintMsg} = 1;
$monitor->{ErrInfo}{SaveMessages} = 1;
$inputbuffer = $monitor->sql("DBCC INPUTBUFFER($spid) WITH NO_INFOMSGS",
                             SINGLEROW, HASH);
$inputbuffer = $monitor->sql("DBCC INPUTBUFFER($spid) WITH NO_INFOMSGS",
                             SINGLESET, HASH);
if ($$inputbuffer[0] and $$inputbuffer[0]{$colname} =~ /^PRINT 'Elfmeter'/) {
   print "ok 12\n";
}
else {
   print "not ok 12 # $$inputbuffer[0]{'EventInfo'}\n";
}
$monitor->{ErrInfo}{PrintText} = 0;

# Testing Appname. There is a default which should be the script name,
$testc = setup_testc;
$testc->{AutoConnect} = 1;
my $name = $testc->sql_one('SELECT app_name()', SCALAR);
if ($name eq '9_loginproperties.t') {
   print "ok 13\n";
}
elsif ($monitorsqlver == 6 and $name eq substr('9_loginproperties.t', 0, 15)) {
   print "ok 13\n";
}
else {
   print "not ok 13 # $name\n";
}

# And set explicit.
$testc->setloginproperty('Appname', 'Papperstapet');
$name = $testc->sql_one('SELECT app_name()', SCALAR);
if ($name eq 'Papperstapet') {
   print "ok 14\n";
}
else {
   print "not ok 14 # $name\n";
}

# Test setting language. This does not work well on 6.5.
if ($monitorsqlver > 6) {
   $testc->setloginproperty('Language', 'Spanish');
   $name = $testc->sql_one("SELECT convert(varchar, convert(datetime, '20030112'))");
   if ($name =~ /^Ene 12 2003/) {
      print "ok 15\n";
   }
   else {
      print "not ok 15 # $name\n";
   }
}
else {
   print "ok 15 # skip\n";
}


# AttachFilename. Only for SQL 2000 and later.
if ($monitorsqlver >= 8) {
   $monitor->{ErrInfo}{PrintText} = 1;  # Suppress output from CREATE/DROP Database
   $monitor->sql('CREATE DATABASE OlleDB$test');
   my @helpdb = $monitor->sql_sp('sp_helpdb', ['OlleDB$test']);
   $monitor->sql_sp('sp_detach_db', ['OlleDB$test']);
   my $filename = $helpdb[1]{'filename'};
   $filename =~ s!\\\\!\\!g;
   $filename =~ s!\s+$!!g;
   $testc = setup_testc;
   $testc->setloginproperty('AttachFilename', $filename);
   $testc->setloginproperty('Database', 'OlleDB test');
   $testc->setloginproperty('Pooling', 0);
   $testc->connect;
   $db = $testc->sql_one('SELECT db_name()');
   if ($db eq 'OlleDB test') {
      print "ok 16\n";
   }
   else {
      print "not ok 16 # $db\n";
   }
   $testc->disconnect;
   $monitor->sql('DROP DATABASE [OlleDB test]');
   $monitor->{ErrInfo}{PrintText} = 0;
}
else {
   print "ok 16 # skip\n";
}

# Network address. This works like server - maybe.
$testc = new Win32::SqlServer;
$testc->{Provider} = $provider if defined $provider;
$testc->setloginproperty('Networkaddress', $mainserver);
if ($mainuser) {
   $testc->setloginproperty('Username', $mainuser);
   $testc->setloginproperty('Password', $mainpw);
}
else {
   $testc->setloginproperty('IntegratedSecurity', "SSPI");
}
$testc->{ErrInfo}{PrintMsg}    = 17;
$testc->{ErrInfo}{PrintLines}  = 17;
$testc->{ErrInfo}{PrintText}   = 17;
$testc->{ErrInfo}{MaxSeverity} = 17;
$testc->{ErrInfo}{SaveMessages} = 1;
$testc->connect;
if (not exists($testc->{ErrInfo}{Messages})) {
   my $servername1 = $monitor->sql_one('SELECT @@servername', SCALAR);
   my $servername2 = $testc->sql_one('SELECT @@servername', SCALAR);
   if ($servername1 eq $servername2) {
      print "ok 17\n";
   }
   else {
      print "not ok 17 # servername = $servername2\n";
   }
}
else {
   print "not ok 17 # " . $testc->{ErrInfo}{Messages}[0]{'text'} . "\n";
}
$testc->disconnect;

# Network Library. We don't test this, because we cannot easily determine
# which protocols the server is running. (xp_regread would do it, but we
# are not being backwards for this.
if (1 == 0) {
   $testc = setup_testc;
   $testc->setloginproperty('NetLib', 'DBNMPNTW');
   $testc->connect;
   my $netlib = $testc->sql_one(<<'SQLEND', SCALAR);
   SELECT net_library FROM master.dbo.sysprocesses WHERE spid = @@spid
SQLEND
   warn "$netlib\n";
}

# Packet size. Only testable on SQL 2005.
if ($monitorsqlver >= 9) {
   $testc = setup_testc;
   $testc->setloginproperty('PacketSize', 1280);
   $testc->connect;
   my $pktsize = $testc->sql_one(<<'SQLEND', SCALAR);
   SELECT net_packet_size FROM sys.dm_exec_connections WHERE session_id = @@spid
SQLEND
   if ($pktsize == 1280) {
      print "ok 18\n";
   }
   else {
      print "not ok 18 # $pktsize\n";
   }
}
else {
   print "ok 18 # skip\n";
}

# Hostname. First test default, then to set name explicitly.
$testc = setup_testc;
$testc->{AutoConnect} = 1;
$name = $testc->sql_one('SELECT host_name()', SCALAR);
if ($name eq Win32::NodeName) {
   print "ok 19\n";
}
else {
   print "not ok 19 # $name\n";
}

# And set explicit.
$testc->setloginproperty('hOsTnAmE', 'Nyckelpiga');
$name = $testc->sql_one('SELECT host_name()', SCALAR);
if ($name eq 'Nyckelpiga') {
   print "ok 20\n";
}
else {
   print "not ok 20 # $name\n";
}

# Test connection string. If this attributes, all other defaults should
# be lost.
$testc = setup_testc;
my $connectstring = "Database=tempdb;";
$connectstring .= "Server=$mainserver;" if $mainserver;
if ($mainuser) {
    $connectstring .= "UID=$mainuser;";
}
if ($mainpw) {
    $connectstring .= "PWD=$mainpw;"
}
if (not ($mainuser or $mainpw)) {
   $connectstring .= "Trusted_connection=Yes;";
}
$connectstring =~ s/;$//;
my $nothostname = (Win32::NodeName ne 'Sture' ? 'Sture' : 'Sten');
$testc->setloginproperty('Hostname', $nothostname);
$testc->setloginproperty('ConnectionString', $connectstring);
$testc->connect;
$name = $testc->sql_one('SELECT host_name()', SCALAR);
if ($name ne $nothostname) {
   print "ok 21\n";
}
else {
   print "not ok 21 # $name\n";
}
$name = $testc->sql_one('SELECT app_name()', SCALAR);
if ($name !~ /^9_login/) {
   print "ok 22\n";
}
else {
   print "not ok 22 # $name\n";
}
$testc->disconnect;

# Test old password. This requires SQL 2005, SQL Native Client and SQL
# authentication.
if ($monitor->{Provider} == PROVIDER_SQLNCLI and $monitorsqlver >= 9 and
    $loginconfig{'config_value'} !~ /^Windows/) {
   my $testuser = 'Olle' . rand;
   my $pw1 = 'pw1' . rand;
   my $pw2 = 'pw2' . rand;
   $monitor->sql("CREATE LOGIN [$testuser] WITH password = '$pw1' ," .
                 "CHECK_POLICY = OFF");

   # Test password change. We run without pooling, to make it possible to
   # drop login at end.
   $testc = setup_testc;
   $testc->setloginproperty('Username', $testuser);
   $testc->setloginproperty('Password', $pw2);
   $testc->setloginproperty('OldPassword', $pw1);
   $testc->setloginproperty('Pooling', 0);
   $testc->connect;
   if (not $testc->{ErrInfo}{Messages}) {
       print "ok 23\n";
   }
   else {
       print "not ok 23\n";
   }
   $testc->disconnect();

   # And test that password really changed.
   $testc = setup_testc;
   $testc->setloginproperty('Username', $testuser);
   $testc->setloginproperty('Password', $pw2);
   $testc->setloginproperty('Pooling', 0);
   $testc->connect;
   if (not $testc->{ErrInfo}{Messages}) {
       print "ok 24\n";
   }
   else {
       print "not ok 24\n";
   }
   $testc->disconnect();

   # Clean up
   $monitor->sql("DROP LOGIN [$testuser]");
}
else {
   print "ok 23 # skip\n";
   print "ok 24 # skip\n";
}

# Test is_connected
$testc = setup_testc;
if (not $testc->isconnected()) {
   print "ok 25\n";
}
else {
   print "not ok 25\n";
}

$testc->connect();
if ($testc->isconnected()) {
   print "ok 26\n";
}
else {
   print "not ok 26\n";
}

$testc->cancelbatch();
if ($testc->isconnected()) {
   print "ok 27\n";
}
else {
   print "not ok 27\n";
}

$testc->disconnect();
if (not $testc->isconnected()) {
   print "ok 28\n";
}
else {
   print "not ok 28\n";
}

# Don't pool this command, SQL 7 has problem with reusing connection
$testc->setloginproperty('Pooling', 0);
$testc->connect();
$testc->{ErrInfo}{MaxSeverity} = 25;
$testc->{ErrInfo}{PrintLines} = 25;
$testc->{ErrInfo}{PrintMsg} = 25;
$testc->{ErrInfo}{PrintText} = 25;
$testc->{ErrInfo}{CarpLevel} = 25;
$testc->sql("RAISERROR('Testing Win32::SqlServer', 20, 1) WITH LOG");
if (not $testc->isconnected()) {
   print "ok 29\n";
}
else {
   print "not ok 29\n";
}

# Finally test DisconnectOn in ErrInfo.
$testc->setloginproperty('Pooling', 1);
$testc->connect();
$testc->sql("SELECT * FROM #nosuchtable");
if ($testc->isconnected()) {
   print "ok 30\n";
}
else {
   print "not ok 30\n";
}

$testc->{CommandTimeout} = 1;
$testc->sql("WAITFOR DELAY '00:00:05'");
if ($testc->isconnected()) {
   print "ok 31\n";
}
else {
   print "not ok 31\n";
}


$testc->{ErrInfo}{DisconnectOn}{'208'}++;
$testc->sql("SELECT * FROM #nosuchtable");
if (not $testc->isconnected()) {
   print "ok 32\n";
}
else {
   print "not ok 32\n";
}

$testc->connect();
$testc->{CommandTimeout} = 1;
$testc->{ErrInfo}{DisconnectOn}{'HYT00'}++;
$testc->sql("WAITFOR DELAY '00:00:05'");
if (not $testc->isconnected()) {
   print "ok 33\n";
}
else {
   print "not ok 33\n";
}

