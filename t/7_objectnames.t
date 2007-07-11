#---------------------------------------------------------------------
# $Header: /Perl/OlleDB/t/7_objectnames.t 5     07-06-10 21:32 Sommar $
#
# This test suite tests that we interpret object names passed to sql_sp
# and sql_insert correctly.
#
# $History: 7_objectnames.t $
# 
# *****************  Version 5  *****************
# User: Sommar       Date: 07-06-10   Time: 21:32
# Updated in $/Perl/OlleDB/t
# Don't use sp_addgroup to create a schema on SQL 2005 or higher, since
# there is CREATE SCHEMA - and in Katmai there is no sp_addgroup.
#
# *****************  Version 4  *****************
# User: Sommar       Date: 05-11-26   Time: 23:47
# Updated in $/Perl/OlleDB/t
# Renamed the module from MSSQL::OlleDB to Win32::SqlServer.
#
# *****************  Version 3  *****************
# User: Sommar       Date: 05-10-30   Time: 22:34
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 2  *****************
# User: Sommar       Date: 05-03-28   Time: 20:01
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 1  *****************
# User: Sommar       Date: 05-03-28   Time: 19:03
# Created in $/Perl/OlleDB/t
#---------------------------------------------------------------------

use strict;
use Win32::SqlServer qw(:DEFAULT :consts);
use File::Basename qw(dirname);

require &dirname($0) . '\testsqllogin.pl';

use vars qw(@testres $verbose $no_of_tests);

sub blurb{
    push (@testres, "#------ Testing @_ ------\n");
    print "#------ Testing @_ ------\n" if $verbose;
}

$verbose = shift @ARGV;

$^W = 1;

$| = 1;

my $X = testsqllogin();
my ($sqlver) = split(/\./, $X->{SQL_version});
my ($sqlncli) = ($X->{Provider} == PROVIDER_SQLNCLI);

# Suppress informatiomal messages for our coming creation craze.
$X->{errInfo}{printText} = 1;

# Permit us to continue on errors.
$X->{ErrInfo}{MaxSeverity} = 17;

# The out data from the test procedures is a return value, so turn off that
# test.
$X->{ErrInfo}{CheckRetStat} = 0;


# This becomes "räksmörgås" - but in Greek script.
my $shrimp = "\x{03A1}\x{03B5}\x{03BA}\x{03C3}\x{03BC}\x{03BF}\x{03B5}\x{03C1}\x{03BD}\x{03B3}\x{03C9}\x{03C2}";

# Database names we use. They are some absymal to avoid collisions with existing
# databases. Names with embedded dots does not work on 6.5, although in theory
# they should.
my @dbs = ('Olle$DB', '"Olle$DB test"');
push(@dbs, '"OlleDB.test"', '[Olle$DB.test]', $shrimp) if $sqlver > 6;

# Schema names that we use. On 6.5, we only test the dbo schema, in SQL 6.5
# groups does not have schemas, and we don't want to create logins to create users.
# Also, users cannot have "funny" characters in them on 6.5.
my @schemas = ('dbo', 'guest');
push(@schemas, '"OlleDB$ test"', '"."', '"OlleDB.."""',
               '[Olle$DB.test]', '[".]', $shrimp) if $sqlver > 6;

# And procedure names.
my @procnames = ('plain_sp', '"space sp"');
push (@procnames, '"dot.sp"', '"dot.dot.sp"', '[bracket sp.]', '[bracket]]sp]', $shrimp) if $sqlver > 6;

# Drop existing databases. This is commented out normally as a safety
# precaution, so that we don't drop existing databases.
#$X->sql("USE master");
#foreach my $db (@dbs) {
#   $X->sql("IF db_id('$db') IS NOT NULL DROP DATABASE $db");
#}


# Go on and create databases, schemas and procedures. Note that we don't drop
# existing databases. If the script fails, you may have drop to the databases
# manually.
my (%procmap, $n);
foreach my $db (@dbs) {
   $X->sql("USE master");
   $X->sql("CREATE DATABASE $db");
   $X->sql("USE $db");

   # Add guest so we can SETUSER to it.
   $X->sql("EXEC sp_adduser guest");

   # And create the schemas as groups (so logins are not required).
   foreach my $sch (@schemas) {
      unless ($sch =~ /^(dbo|guest)$/) {
         if ($sqlver >= 9) {
            $X->sql("CREATE SCHEMA $sch");
         }
         else {
            # No direct CREATE SCHEMA in previous version, but creating a
            # group will do.
            $X->sql("EXEC sp_addgroup $sch");
         }
      }

      # And so the procedures. Each procedure has a unique signature with
      # the parameter name, and we save this in %procmap.
      foreach my $proc (@procnames) {
         $n++;
         $X->sql ("CREATE PROCEDURE $sch.$proc \@a$n int AS RETURN \@a$n + $n");
         $X->sql ("GRANT EXECUTE ON $sch.$proc TO public");
         $procmap{$db}{$sch}{$proc} = $n;

         # On SQL 2005 and later, also create schema collections to test
         # handling of typeinfo if we have SQL Native client.
         if ($sqlver >= 9 and $sqlncli) {
            $X->sql(<<SQLEND);
CREATE XML SCHEMA COLLECTION $sch.$proc AS '
<schema xmlns="http://www.w3.org/2001/XMLSchema">
      <element name="Olle$n" type="string"/>
</schema>'
SQLEND
         }
      }
   }
}

# Also create a temporary stored procedure.
$X->sql('CREATE PROCEDURE #temp_sp @a4711 int AS RETURN 14711');
if ($sqlver >= 9  and $sqlncli) {
  $X->sql(<<SQLEND);
CREATE XML SCHEMA COLLECTION #temp_sp AS
'<schema xmlns="http://www.w3.org/2001/XMLSchema">
      <element name="Olle4711" type="string"/>
</schema>'
SQLEND
}

# First try all SP without schema qualification in the first database.
my $db = $dbs[0];
$X->sql("USE $db");
my $sch = 'dbo';
foreach my $proc (@procnames) {
   my $expect = $procmap{$db}{$sch}{$proc};
   do_test($proc, $expect, 1);
   do_test(".$proc", $expect);
   do_test("..$proc", $expect);
   do_test("...$proc", $expect);
   do_test("$db. .$proc", $expect, 1);
   do_test("....$proc", 'ERROR');
   do_test(".$db.$sch.$proc", $expect);
   do_test(". $db . $sch . $proc", $expect);
   do_test(". $db . $sch . $proc", $expect);  # Do it twice to test look-up.
   do_test("...$sch.$proc", 'ERROR');
   do_test("server.$db.$sch.$proc", 'ERROR', 1);
   do_test("..$db.$sch.$proc", 'ERROR');
}

# Redo for the guest schema. We must flush the proc cache here.
$X->sql("SETUSER 'guest'");
$X->{'procs'} = {};
$sch = 'guest';
foreach my $proc (@procnames) {
   my $expect = $procmap{$db}{$sch}{$proc};
   do_test($proc, $expect, 1);
   do_test("guest.$proc", $expect, 1);
   do_test("..$proc", $expect);
   do_test(". ..$proc", $expect);
}
$X->sql("SETUSER");

# Now try all combinations of schema and procedure.
foreach $sch (@schemas) {
   foreach my $proc (@procnames) {
      my $expect = $procmap{$db}{$sch}{$proc};
      do_test(" $sch.$proc ", $expect, 1);
      do_test(".$sch.$proc", $expect);
      do_test("..$sch.$proc", $expect);
   }
}

# And now all combinations of databases, schemas and procedeurs.
$X->sql("USE master");
foreach $db (@dbs) {
   foreach $sch (@schemas) {
      foreach my $proc (@procnames) {
         my $expect = $procmap{$db}{$sch}{$proc};
         do_test("$db.$sch.$proc", $expect, 1);
      }
   }
}

# Test the temporary stored procedure.
blurb("#temp_sp");
do_test("#temp_sp", 4711);

# Finnaly test system stored procedures.
my $resset = ($sqlver ==  6 ? 2 : 1);
$X->sql("USE $dbs[0]");
blurb("sp_help plain_sp");
my @result = sql_sp('sp_help', ['plain_sp']);
push(@testres,
     $result[$resset]{'Parameter_name'} eq '@a' . $procmap{$dbs[0]}{'dbo'}{'plain_sp'});
foreach $db (@dbs) {
   blurb("$db..sp_help plain_sp");
   @result = sql_sp("$db..sp_help", ['plain_sp']);
   push(@testres,
         $result[$resset]{'Parameter_name'} eq '@a' . $procmap{$db}{'dbo'}{'plain_sp'});
}


$X->sql("USE master");
foreach my $db (@dbs) {
   $X->sql("DROP DATABASE $db");
}

if ($sqlver >= 9 and $sqlncli) {
   $no_of_tests = 917;
}
elsif ($sqlver > 6) {
   $no_of_tests =  553;
}
else {
   $no_of_tests = 52;
}

finally:



my $ix = 1;
my $blurb = "";
print "1..$no_of_tests\n";
foreach my $result (@testres) {
   if ($result =~ /^#--/) {
      print $result if $verbose;
      $blurb = $result;
   }
   elsif ($result == 1) {
      printf "ok %d\n", $ix++;
   }
   else {
      printf "not ok %d\n$blurb", $ix++;
   }
}

exit;

sub do_test {
    my($objref, $mapvalue, $doxml) = @_;
    blurb($objref);
    if ($mapvalue =~ /^\d+$/) {
       my $retvalue;
       my $params;
       $$params{"a$mapvalue"} = 10000;
       my $expect = 10000 + $mapvalue;
       $X->sql_sp($objref, \$retvalue, $params);
       push(@testres, $retvalue == $expect);

       if ($sqlver >= 9 and $sqlncli and $doxml) {
          $expect = "<Olle$mapvalue>$mapvalue</Olle$mapvalue>";
          my $sqlparams = ['xml', '<?xml version="1.0"?>' . $expect, $objref];
          $retvalue = $X->sql('SELECT convert(nvarchar(MAX), ?)', [$sqlparams],
                               SCALAR, SINGLEROW);
          push(@testres, $retvalue eq $expect);
       }
    }
    else {
       delete $X->{ErrInfo}{Messages};
       $X->{ErrInfo}{PrintMsg} = 17;
       $X->{ErrInfo}{PrintLines} = 17;
       $X->{ErrInfo}{PrintText} = 17;
       $X->{ErrInfo}{CarpLevel} = 17;
       $X->{ErrInfo}{SaveMessages} = 1;
       $X->sql_sp($objref);
       my $errmsg = $X->{ErrInfo}{Messages}[0]{'text'};
       push(@testres, $errmsg =~ /Stored procedure .* not accessible/);

       if ($sqlver >= 9 and $sqlncli and $doxml) {
          delete $X->{ErrInfo}{Messages};
          $X->sql('SELECT ?', ['xml', undef, $objref]);
          $errmsg = $X->{ErrInfo}{Messages}[0]{'text'};
          push(@testres, $errmsg =~ /Incorrect syntax near/);
       }

       $X->{ErrInfo}{PrintMsg} = 1;
       $X->{ErrInfo}{PrintLines} = 11;
       $X->{ErrInfo}{PrintText} = 1;
       $X->{ErrInfo}{CarpLevel} = 10;
       $X->{ErrInfo}{SaveMessages} = 0;
    }
}

