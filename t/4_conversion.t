#---------------------------------------------------------------------
# $Header: /Perl/OlleDB/t/4_conversion.t 2     05-11-26 23:47 Sommar $
#
# Tests that it's possible to set up a conversion based on the local
# OEM character set and the server charset. Mainly is this is test that
# we can access Win32::Registry properly.
#
# $History: 4_conversion.t $
# 
# *****************  Version 2  *****************
# User: Sommar       Date: 05-11-26   Time: 23:47
# Updated in $/Perl/OlleDB/t
# Renamed the module from MSSQL::OlleDB to Win32::SqlServer.
#
# *****************  Version 1  *****************
# User: Sommar       Date: 05-02-06   Time: 22:51
# Created in $/Perl/OlleDB/t
#---------------------------------------------------------------------

use strict;
use Win32::SqlServer qw(:DEFAULT :consts);
use File::Basename qw(dirname);

require &dirname($0) . '\testsqllogin.pl';

$^W = 1;
$| = 1;


my($shrimp, $shrimp_850, $shrimp_twoway, $shrimp_bogus, @data, $data, %data);

# Get client char-set.
my $client_cs = get_codepage_from_reg('OEMCP');

# These are the constants we use to test. It's all about shrimp sandwiches.
$shrimp       = 'räksmörgås';  # The way it should be in Latin-1.
if ($client_cs == 850) {
   $shrimp_850    = 'r„ksm”rg†s';  # It's in CP850.
   $shrimp_twoway = 'räksmörgås';  # Latin-1 -> CP850 and back.
   $shrimp_bogus  = 'rõksm÷rgÕs';  # Converted to Latin-1 as if it was CP850 but it wasn't.
}
elsif ($client_cs == 437) {
   $shrimp_850    = 'r„ksm”rg†s';  # It's in CP437.
   $shrimp_twoway = 'r_ksmörg_s';  # Latin-1 -> Cp437 and back. Not round-trip.
   $shrimp_bogus  = 'r_ksm÷rg_s';  # Converted to Latin-1 as if it was CP437 but it wasn't.
}
else {
   print "Skipping this test; no test defined for code-page $client_cs\n";
   print "1..0\n";
   exit;
}

print "1..23\n";

my $X = testsqllogin();

# First create a table to two procedures to read and write to a table.
sql(<<SQLEND);
   CREATE TABLE #nisse (i       int      NOT NULL PRIMARY KEY,
                        shrimp  char(10) NOT NULL)
SQLEND

sql(<<'SQLEND');
   CREATE PROCEDURE #nisse_ins_sp @i      int,
                                  @shrimp char(10) AS
      INSERT #nisse (i, shrimp) VALUES (@i, @shrimp)
SQLEND

sql(<<'SQLEND');
   CREATE PROCEDURE #nisse_get_sp @i int,
                                  @shrimp char(10) OUTPUT AS

      SELECT @shrimp = shrimp FROM #nisse WHERE @i = i
SQLEND


# Now add first set of data with no conversion in effect..
sql("INSERT #nisse (i, shrimp) VALUES (0, 'räksmörgås')");
sql("INSERT #nisse (i, shrimp) VALUES (?, ?)", [['int', 1], ['char', 'räksmörgås']]);
sql_insert("#nisse", {i => 2, 'shrimp' => 'räksmörgås'});
sql_sp("#nisse_ins_sp", [3, 'räksmörgås']);

# Now set up default, bilateral conversion.
sql_set_conversion;
print "ok 1\n";   # We wouldn't come back if it's not ok...

# Add a second set of data, now conversion is in effect.
sql("INSERT #nisse (i, shrimp) VALUES (10, 'räksmörgås')");
sql("INSERT #nisse (i, shrimp) VALUES (?, ?)", [['int', 11], ['char', 'räksmörgås']]);
sql_insert("#nisse", {i => 12, 'shrimp' => 'räksmörgås'});
sql_sp("#nisse_ins_sp", [13, 'räksmörgås']);

# Now retrieve data and see what we get. The first should give the shrimp in CP850.
@data = sql("SELECT shrimp FROM #nisse WHERE i BETWEEN 0 AND 3", SCALAR);
if (compare(\@data, [$shrimp_850, $shrimp_850, $shrimp_850, $shrimp_850])) {
   print "ok 2\n";
}
else {
   print "not ok 2\n# " . join(' ', @data) . "\n";
}

# This should give the real McCoy - it's been converted in both directions.
@data = sql("SELECT shrimp FROM #nisse WHERE i BETWEEN 10 AND 13", SCALAR);
if (compare(\@data,
            [$shrimp_twoway, $shrimp_twoway, $shrimp_twoway, $shrimp_twoway])) {
   print "ok 3\n";
}
else {
   print "not ok 3\n# " . join(' ', @data) . "\n";
}

# Again, a CP850 shrimp is expected.
sql_sp("#nisse_get_sp", [1, \$data]);
if ($data eq $shrimp_850) {
   print "ok 4\n";
}
else {
   print "not ok 4\n# $data\n";
}

# Again, in Latin-1.
sql_sp("#nisse_get_sp", [11, \$data]);
if ($data eq $shrimp_twoway) {
   print "ok 5\n";
}
else {
   print "not ok 5\n# $data\n";
}

# Turn off conversion. This just can't fail. :-)
sql_unset_conversion;

# Now we should get Latin-1.
@data = sql("SELECT shrimp FROM #nisse WHERE i BETWEEN 0 AND 3", SCALAR);
if (compare(\@data, [$shrimp, $shrimp, $shrimp, $shrimp])) {
   print "ok 6\n";
}
else {
   print "not ok 6\n# " . join(' ', @data) . "\n";
}

# This is the bogus conversion, we converted Latin-1 to Latin-1.
@data = sql("SELECT shrimp FROM #nisse WHERE i BETWEEN 10 AND 13", SCALAR);
if (compare(\@data,
             [$shrimp_bogus, $shrimp_bogus, $shrimp_bogus, $shrimp_bogus])) {
   print "ok 7\n";
}
else {
   print "not ok 7\n# " . join(' ', @data) . "\n";
}

# Again, a Latin-1 shrimp is expected.
sql_sp("#nisse_get_sp", [1, \$data]);
if ($data eq $shrimp) {
   print "ok 8\n";
}
else {
   print "not ok 8\n# $data\n";
}

# Again, it's bogus.
sql_sp("#nisse_get_sp", [11, \$data]);
if ($data eq $shrimp_bogus) {
   print "ok 9\n";
}
else {
   print "not ok 9\n# $data\n";
}


# Now we will make a test that we convert hash keys correctly. We will also
# test asymmetric conversion and that sql_one converts properly.
sql_set_conversion("CP$client_cs", "iso_1", TO_CLIENT_ONLY);
{
   my %ref;
   $ref{$shrimp_850} = $shrimp_850;

   %data = sql(q!SELECT "räksmörgås" = 'räksmörgås'!, HASH, SINGLEROW);
   if (compare(\%ref, \%data)) {
      print "ok 10\n";
   }
   else {
      print "not ok 10\n";
   }

   %data = sql_one(q!SELECT "räksmörgås" = 'räksmörgås'!);
   if (compare(\%ref, \%data)) {
      print "ok 11\n";
   }
   else {
      print "not ok 11\n";
   }
}

# After this we have conversion both directions
sql_set_conversion("CP$client_cs", "iso_1", TO_SERVER_ONLY);
{
   my %ref;
   $ref{$shrimp_twoway} = $shrimp_twoway;

   %data = sql("SELECT 'räksmörgås' = 'räksmörgås'", HASH, SINGLEROW);
   if (compare(\%ref, \%data)) {
      print "ok 12\n";
   }
   else {
      print "not ok 12\n";
   }

   %data = sql_one("SELECT 'räksmörgås' = 'räksmörgås'");
   if (compare(\%ref, \%data)) {
      print "ok 13\n";
   }
   else {
      print "not ok 13\n";
   }
}

# After now only to server.
sql_unset_conversion(TO_CLIENT_ONLY);
{
   my %ref;
   $ref{$shrimp_bogus} = $shrimp_bogus;

   %data = sql(q!SELECT "räksmörgås" = 'räksmörgås'!, HASH, SINGLEROW);
   if (compare(\%ref, \%data)) {
      print "ok 14\n";
   }
   else {
      print "not ok 14\n";
      print '<' . (keys(%ref))[0] . '> <' . (keys(%data))[0] . ">\n";
   }

   %data = sql_one(q!SELECT "räksmörgås" = 'räksmörgås'!);
   if (compare(\%ref, \%data)) {
      print "ok 15\n";
   }
   else {
      print "not ok 15\n";
   }
}

# And now in no direction at all.
sql_unset_conversion(TO_SERVER_ONLY);
{
   my %ref;
   $ref{$shrimp} = $shrimp;

   %data = sql(q!SELECT "räksmörgås" = 'räksmörgås'!, HASH, SINGLEROW);
   if (compare(\%ref, \%data)) {
      print "ok 16\n";
   }
   else {
      print "not ok 16\n";
   }

   %data = sql_one(q!SELECT "räksmörgås" = 'räksmörgås'!);
   if (compare(\%ref, \%data)) {
      print "ok 17\n";
   }
   else {
      print "not ok 17\n";
   }
}

if ($client_cs == 850) {
   # Now we will test with object name that are subject to conversion. First
   # some tables. Turn off conversion before anything else!
   # This test only works with CP850, as CP437 is not roundtrip.
   sql_unset_conversion;
   sql(<<SQLEND);
      CREATE TABLE #$shrimp (i       int     NOT NULL PRIMARY KEY,
                            $shrimp  char(9) NOT NULL)
SQLEND

   sql(<<SQLEND);
      CREATE PROCEDURE #${shrimp}_ins_sp \@i       int,
                                         \@$shrimp char(9) AS
         INSERT #$shrimp (i, $shrimp) VALUES (\@i, \@$shrimp)
SQLEND

   sql(<<SQLEND);
      CREATE PROCEDURE #${shrimp}_get_sp \@i int,
                                         \@$shrimp char(9) OUTPUT AS

         SELECT \@$shrimp = $shrimp FROM #$shrimp WHERE \@i = i
SQLEND

   # Insert some data
   sql("INSERT #$shrimp (i, $shrimp) VALUES (1, 'first row')");
   if ($X->{SQL_version} =~ /^6\./) {
      sql("INSERT #$shrimp (i, $shrimp) VALUES (?, ?)",
          [['int', 2], ['char', 'secondrow']]);
   }
   else {
      sql("INSERT #$shrimp (i, $shrimp) VALUES (\@i, \@$shrimp)",
          {i => ['int', 2], $shrimp => ['char', 'secondrow']});
   }
   sql_insert("#$shrimp", {i => 3, $shrimp => 'third row'});
   sql_sp("#${shrimp}_ins_sp", [4, 'fourthrow']);

   # Turn on conversion.
   sql_set_conversion;

   # We assume that things just crashes if test fails.
   sql("INSERT #$shrimp_850 (i, $shrimp_850) VALUES (5, 'fifth row')");
   print "ok 18\n";
   if ($X->{SQL_version} =~ /^6\./) {
      sql("INSERT #$shrimp_850 (i, $shrimp_850) VALUES (?, ?)",
          [['int', 6], ['char', 'sixth row']]);
   }
   else {
      sql("INSERT #$shrimp_850 (i, $shrimp_850) VALUES (\@i, \@$shrimp_850)",
          {i => ['int', 6], $shrimp_850 => ['char', 'sixth row']});
   }
   print "ok 19\n";
   sql_insert("#$shrimp_850", {i => 7, $shrimp_850 => 'row seven'});
   print "ok 20\n";
   sql_sp("#${shrimp_850}_ins_sp", [8, 'eighthrow']);
   print "ok 21\n";

   # Check that data was inserted as expected.
   @data = sql("SELECT $shrimp_850 FROM #$shrimp_850 ORDER BY i", SCALAR);
   if (compare(\@data, ['first row', 'secondrow', 'third row', 'fourthrow',
                        'fifth row', 'sixth row', 'row seven', 'eighthrow'])) {
      print "ok 22\n";
   }
   else {
      print "not ok 22\n# " . join(' ', @data) . "\n";
   }
}
else {
   print "ok 18 # skip, test cannot work on CP437\n";
   print "ok 19 # skip, test cannot work on CP437\n";
   print "ok 20 # skip, test cannot work on CP437\n";
   print "ok 21 # skip, test cannot work on CP437\n";
   print "ok 22 # skip, test cannot work on CP437\n";
}

# Final test: check that a datetime hash is not thrashed when subject to
# conversion
sql_set_conversion;
$X->{DatetimeOption} = DATETIME_HASH;
$data = sql_one('SELECT dateadd(YEAR, 100, dateadd(minute, 20, ?))',
                [['datetime', '18140212 17:19:34']], SCALAR);
if (compare($data, {Year => 1914, Month => 2, Day => 12,
                    Hour => 17, Minute => 39, Second => 34, Fraction => 0})) {
   print "ok 23\n";
}
else {
   print "not ok 23\n";
}

exit;

sub compare {
   my ($x, $y) = @_;

   my ($refx, $refy, $ix, $key, $result);

   $refx = ref $x;
   $refy = ref $y;

   if (not $refx and not $refy) {
      if (defined $x and defined $y) {
         warn "<$x> ne <$y>" if $x ne $y;
         return ($x eq $y);
      }
      else {
         return (not defined $x and not defined $y);
      }
   }
   elsif ($refx ne $refy) {
      return 0;
   }
   elsif ($refx eq "ARRAY") {
      if ($#$x != $#$y) {
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
      if ($nokeys_x != $nokeys_y) {
         return 0;
      }
      elsif ($nokeys_x > 0) {
         foreach $key (keys %$x) {
            if (not exists $$y{$key}) {
                return 0;
            }
            $result = compare($$x{$key}, $$y{$key});
            last if not $result;
         }
         return $result;
      }
      else {
         return 1;
      }
   }
   elsif ($refx eq "SCALAR") {
      return compare($$x, $$y);
   }
   else {
      return ($x eq $y);
   }
}

#--------------------------------- Copied from Sqllib.pm
sub get_codepage_from_reg {
    my($cp_value) = shift @_;
    # Reads the code page for OEM or ANSI. This is one specific key in
    # in the registry.

    my($REGKEY) = 'SYSTEM\CurrentControlSet\Control\Nls\CodePage';
    my($regref, $dummy, $result);

    # We need this module to read the registry, but as this is the only
    # place we need it in, we don't C<use> it.
    require 'Win32\Registry.pm';

    $dummy = $main::HKEY_LOCAL_MACHINE;  # Resolve "possible typo" with AS Perl.
    $main::HKEY_LOCAL_MACHINE->Open($REGKEY, $regref) or
         die "Could not open registry key: '$REGKEY'\n";

    # This is where stuff is getting really ugly, as I have found no code
    # that works both with the ActiveState Perl and the native port.
    if ($] < 5.004) {
       Win32::RegQueryValueEx($regref->{'handle'}, $cp_value, 0,
                              $dummy, $result) or
             die "Could not read '$REGKEY\\$cp_value' from registry\n";
    }
    else {
       $regref->QueryValueEx($cp_value, $dummy, $result);
    }
    $regref->Close or warn "Could not close registry key.\n";

    $result;
}
