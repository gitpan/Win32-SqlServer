#---------------------------------------------------------------------
# $Header: /Perl/OlleDB/t/6_paramsql.t 8     07-06-10 21:50 Sommar $
#
# This test suite concerns sql with parameterised SQL statements.
#
# $History: 6_paramsql.t $
# 
# *****************  Version 8  *****************
# User: Sommar       Date: 07-06-10   Time: 21:50
# Updated in $/Perl/OlleDB/t
# Corrected for a new error message on Katmai in one case.
#
# *****************  Version 7  *****************
# User: Sommar       Date: 05-11-26   Time: 23:47
# Updated in $/Perl/OlleDB/t
# Renamed the module from MSSQL::OlleDB to Win32::SqlServer.
#
# *****************  Version 6  *****************
# User: Sommar       Date: 05-10-29   Time: 23:18
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 5  *****************
# User: Sommar       Date: 05-08-07   Time: 22:41
# Updated in $/Perl/OlleDB/t
# Test case-insensitivty.
#
# *****************  Version 4  *****************
# User: Sommar       Date: 05-07-25   Time: 0:41
# Updated in $/Perl/OlleDB/t
# Added tests for XML and UDT.
#
# *****************  Version 3  *****************
# User: Sommar       Date: 05-06-26   Time: 22:36
# Updated in $/Perl/OlleDB/t
# Now checks 6.5. Added test for (too) long binary and string values.
#
# *****************  Version 2  *****************
# User: Sommar       Date: 05-03-20   Time: 21:48
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 1  *****************
# User: Sommar       Date: 05-03-20   Time: 21:23
# Created in $/Perl/OlleDB/t
#---------------------------------------------------------------------

use strict;
use Win32::SqlServer qw(:DEFAULT :consts);
use File::Basename qw(dirname);

require &dirname($0) . '\testsqllogin.pl';
require '..\helpers\assemblies.pl';

use vars qw(@testres $verbose $retvalue $no_of_tests);

sub blurb{
    push (@testres, "#------ Testing @_ ------\n");
    print "#------ Testing @_ ------\n" if $verbose;
}

sub datehash_compare {
  # Help routine to compare datehashes.
    my($val, $expect) = @_;

    foreach my $part (keys %$expect) {
       return 0 if not defined $$val{$part} or $$expect{$part} != $$val{$part};
    }
    return 1;
}

$verbose = shift @ARGV;


$^W = 1;

$| = 1;

my $X = testsqllogin();
#open (F, '>paramsql.sql');
#$X->{LogHandle} = \*F;

my ($sqlver) = split(/\./, $X->{SQL_version});

my ($result, $expect);

# Accept all errors, and be silent about them, but save messages.
$X->{errInfo}{maxSeverity}   = 25;
$X->{errInfo}{printLines} = 25;
$X->{errInfo}{printMsg}   = 25;
$X->{errInfo}{printText}  = 25;
$X->{errInfo}{carpLevel}  = 25;
$X->{ErrInfo}{SaveMessages} = 1;

blurb ("datetime HASH in and out");
$X->{DatetimeOption} = DATETIME_HASH;
$expect = [{'d' => {Year => 1945, Month => 5, Day => 9,
                   Hour => 12, Minute => 14, Second => 0, Fraction => 0}}];
$result = $X->sql(
   'SELECT d = dateadd(YEAR, 27, dateadd(MONTH, -6, dateadd(DAY, -2, ?)))',
   [['Smalldatetime', {Year => 1918, Month => 11, Day => 11,
                       Hour => 12, Minute => 14}]]);
push(@testres, compare($expect, $result));

blurb ("decimal with variations");
$X->{DecimalAsStr} = 1;
$expect = [{'d1' => '246246246246',
            'd2' => '3578',
            'd3' => '246246246246.246246',
            'n1' => '94.22',
            'n2' => '94.22'}];
$result = $X->sql('SELECT d1 = 2*?, d2 = 2*?, d3 = 2*?, n1 = 2*?, n2 = 2*?',
                  [['decimal',          '123123123123.123123'],
                   ['decimal(10)',      '1789.44'],
                   ['decimal(18,6)',    '123123123123.123123'],
                   ['numeric(8, 2) ',   '47.11'],
                   ['numeric( 8 , 2 )', '47.11']]);
push(@testres, compare($expect, $result));

blurb("binary as binary");
$X->{BinaryAsStr} = 0;
if ($sqlver > 6) {
   $expect = [{'a' => "ABCDEFABCDEF", 'b' => "ABCDEF\0\0ABCDEF", 'c' => "ABCABC"}];
}
else {
   $expect = [{'a' => "ABCDEFABCDEF", 'b' => "ABCDEFABCDEF", 'c' => "ABCABC"}];
}
$result = $X->sql('SELECT a = ? + ?, b = ? + ?, c = ? + ?',
                  [['binary', 'ABCDEF'],
                   ['binary', 'ABCDEF'],
                   ['binary(8)', 'ABCDEF'],
                   ['varbinary( 8)', 'ABCDEF'],
                   ['binary( 3 )', 'ABCDEF'],
                   ['binary(3 )', 'ABCDEF']]);
push (@testres, compare($expect, $result));

blurb("Testing binary as string");
$X->{BinaryAsStr} = 1;
if ($sqlver > 6) {
   $expect = [{'a' => "ABCDEFABCDEF", 'b' => "ABCDEF0000ABCDEF",
               'c' => "ABCDABCD"}];
}
else {
   $expect = [{'a' => "ABCDEFABCDEF", 'b' => "ABCDEFABCDEF",
               'c' => "ABCDABCD"}];
}
$result = $X->sql('SELECT a = ? + ?, b = ? + ?, c = ? + ?',
                  [['varbinary', 'ABCDEF'],
                   ['binary', 'ABCDEF'],
                   ['binary(5)', 'ABCDEF'],
                   ['varbinary( 5)', 'ABCDEF'],
                   ['binary( 2 )', 'ABCDEF'],
                   ['binary(2 )', 'ABCDEF']]);
push (@testres, compare($expect, $result));

blurb("binary as 0x");
$X->{BinaryAsStr} = 'x';
if ($sqlver > 6) {
   $expect = [{'a' => "0xABCDEFABCDEF", 'b' => "0xABCDEF0000ABCDEF",
               'c' => "0xABCDABCD"}];
}
else {
   $expect = [{'a' => "0xABCDEFABCDEF", 'b' => "0xABCDEFABCDEF",
               'c' => "0xABCDABCD"}];
}
$result = $X->sql('SELECT a = ? + ?, b = ? + ?, c = ? + ?',
                  [['binary', '0xABCDEF'],
                   ['varbinary', '0xABCDEF'],
                   ['binary(5)', '0xABCDEF'],
                   ['varbinary( 5)', '0xABCDEF'],
                   ['binary( 2 )', '0xABCDEF'],
                   ['binary(2 )', '0xABCDEF']]);
push (@testres, compare($expect, $result));

blurb("char/varchar");
if ($sqlver > 6) {
   $expect = {'a' => "x'zx''z ", 'b' => "0xABCDEF  0xABCDEF",
               'c' => "0xAB0xAB"};
}
else {
   $expect = {'a' => "x'zx''z ", 'b' => "0xABCDEF0xABCDEF",
               'c' => "0xAB0xAB"};
}
$result = $X->sql_one('SELECT a = ? + ?, b = ? + ?, c = ? + ?',
                      [['char', "x'z"],
                       ['varchar', "x''z "],
                       ['char(10)', '0xABCDEF'],
                       ['varchar( 10)', '0xABCDEF'],
                       ['char( 4 )', '0xABCDEF'],
                       ['char(4 )', '0xABCDEF']], HASH);
push (@testres, compare($expect, $result));

blurb("many parameters");
{
   my $sqlstring = 'SELECT a0 = 0';
   my @params;
   $expect = [{'a0' => 0}];
   foreach my $i (1..1123) {
      $sqlstring .= ", a$i = ?";
      $$expect[0]{"a$i"} = -$i;
      push(@params, ['int', -$i]);
   }
   $result = $X->sql($sqlstring, \@params);
   push (@testres, compare($expect, $result));
}

blurb("too long varchar");
if ($sqlver == 6) {
   $expect = {'a' => 'HELLO DOLLY! ' x 19 . 'HELLO DO'};
}
elsif ($sqlver <= 8) {
   $expect = {'a' => 'HELLO DOLLY! ' x 615 . 'HELLO'};
}
else {
   $expect = {'a' => 'HELLO DOLLY! ' x 1854};
}
$result = $X->sql_one('SELECT a = upper(?)',
              [['varchar',  'Hello Dolly! ' x 1854]], HASH);
push (@testres, compare($expect, $result));

blurb("too long char");
if ($sqlver == 6) {
   $expect = {'a' => 'HELLO DOLLY! ' x 19 . 'HELLO DO'};
}
else {
   $expect = {'a' => 'HELLO DOLLY! ' x 615 . 'HELLO'};
}
$result = $X->sql_one('SELECT a = upper(?)',
              [['char',  'Hello Dolly! ' x 1854]], HASH);
push (@testres, compare($expect, $result));

blurb("too long binary as str");
$X->{BinaryAsStr} = 1;
if ($sqlver == 6) {
   $expect = {'a' => '961147' . '0201AB60961147' x 36,
              'b' => '47119660AB0102' x 36 . '471196'};
}
elsif ($sqlver <= 8) {
   $expect = {'a' => '01AB60961147' . '0201AB60961147' x 1142,
              'b' => '47119660AB0102' x 1142 . '47119660AB01'};
}
else {
   $expect = {'a' => '0201AB60961147' x 7453,
              'b' => '47119660AB0102' x 1142 . '47119660AB01'};
}
$result = $X->sql_one('SELECT a = convert(image, reverse(?)),
                              b = ? + ?',
              [['varbinary', '47119660AB0102' x 7453],
               ['binary',    '47119660AB0102' x 5453],
               ['binary',    '47119660AB0102' x 5453]],
              HASH);
push (@testres, compare($expect, $result));

blurb("too long binary as 0x");
$X->{BinaryAsStr} = 'x';
if ($sqlver == 6) {
   $expect = {'a' => '0x' . '961147' . '0201AB60961147' x 36,
              'b' => '0x' . '47119660AB0102' x 36 . '471196'};
}
elsif ($sqlver <= 8) {
   $expect = {'a' => '0x' . '01AB60961147' . '0201AB60961147' x 1142,
              'b' => '0x' . '47119660AB0102' x 1142 . '47119660AB01'};
}
else {
   $expect = {'a' => '0x' . '0201AB60961147' x 7453,
              'b' => '0x' . '47119660AB0102' x 1142 . '47119660AB01'};
}
$result = $X->sql_one('SELECT a = convert(image, reverse(?)),
                              b = ? + ?',
              [['varbinary', '0x' . '47119660AB0102' x 7453],
               ['binary',    '0x' . '47119660AB0102' x 5453],
               ['binary',    '0x' . '47119660AB0102' x 5453]],
              HASH);
push (@testres, compare($expect, $result));

blurb("Too long binary as binary");
$X->{BinaryAsStr} = 0;
if ($sqlver == 6) {
   $expect = {'a' => '174' . '2010BA06691174' x 18,
              'b' => '47119660AB0102' x 18 . '471'};
}
elsif ($sqlver <= 8) {
   $expect = {'a' => '691174' . '2010BA06691174' x 571,
              'b' => '47119660AB0102' x 571 . '471196'};
}
else {
   $expect = {'a' => '2010BA06691174' x 7453,
              'b' => '47119660AB0102' x 571 . '471196'};
}
$result = $X->sql_one('SELECT a = convert(image, reverse(?)),
                              b = ? + ?',
              [['varbinary', '47119660AB0102' x 7453],
               ['binary',    '47119660AB0102' x 5453],
               ['BINARY',    '47119660AB0102' x 5453]], HASH);
push (@testres, compare($expect, $result));

$no_of_tests = 12;

if ($sqlver == 6) {
   goto finally;
}

blurb("Not expanding '?' in /* */");
$expect = [{'a' => 12, 'c' => 19}];
$result = $X->sql('SELECT a = ?, /* b = ?, */ c = ?', [['int', 12], ['int', 19]]);
push (@testres, compare($expect, $result));

blurb("Not expanding '?' after --");
$expect = [{'a' => 12, 'c' => 19}];
$result = $X->sql(<<SQLEND,  [['int', 12], ['int', 19]]);
SELECT a = ?, -- b = ?,
       c = ?
SQLEND
push (@testres, compare($expect, $result));

blurb("-- in /*");
$expect = [{'a' => 12, 'Col 2' => 14, 'c' => 19}];
$result = $X->sql(<<SQLEND,  [['int', 12], ['int', 14], ['int', 19]]);
SELECT a = ?, /*
       -- b = */ ?,
       c = ?
SQLEND
push (@testres, compare($expect, $result));

blurb("Nested /*");
$expect = [{'a' => 12, 'Col 2' => 19}];
$result = $X->sql(<<SQLEND,  [['int', 12], ['int', 19]]);
SELECT a = ?, /* b = ?, /*
       c = ?, */ d = */ ?
SQLEND
push (@testres, compare($expect, $result));

blurb("Not expanding '?' literal");
$expect = {'a' => 12, 'b' => '?', 'c' => 19};
$result = $X->sql_one("SELECT a = ?, b = '?', c = ?", [['int', 12], ['int', 19]], HASH);
push (@testres, compare($expect, $result));

blurb("Ignoring '/*' in literal");
$expect = [{'a' => 12, 'b' => '/*', 'c' => 19}];
$result = $X->sql(<<SQLEND,  [['int', 12], ['int', 19]]);
SELECT a = ?, b = '/*', c = ?
SQLEND
push (@testres, compare($expect, $result));

blurb("Not expanding '?' in quoted identifiers");
$expect = [{'a?' => 12, '?' => 456, 'c' => 19}];
$result = $X->sql(<<SQLEND,  [['int', 12], ['int', 456], ['int', 19]]);
SELECT "a?" = ?, [?] = ?, c = ?
SQLEND
push (@testres, compare($expect, $result));

blurb("doubling of quotes");
$expect = {'a' => 12, 'b"?' => "?'?987", 'c' => 19};
$result = $X->sql_one(<<SQLEND,  [['int', 12], ['int', 987], ['int', 19]], HASH);
SELECT a = ?, "b""?" = '?''?' + ltrim(str(?)), c = ?
SQLEND
push (@testres, compare($expect, $result));

blurb("doubling of brackets");
$expect = [{'a' => 12, 'b]?' => "?[?987", 'c' => 19}];
$result = $X->sql(<<SQLEND,  [['int', 12], ['int', 987], ['int', 19]]);
SELECT a = ?, [b]]?] = '?[?' + ltrim(str(?)), c = ?
SQLEND
push (@testres, compare($expect, $result));


blurb("expansion of ???");
$expect = [{'a' => 12, 'c' => 19}];
$result = $X->sql("SELECT a = ???,  c = ?",
                  {'@P1@P2@P3' => ['int', 12], '@P4' => ['int', 19]});
push (@testres, compare($expect, $result));

blurb("Expansion of ??? at end of string");
$expect = {'a' => 12, 'c' => 19};
$result = $X->sql_one("SELECT a = ?,  c = ???",
                  {'@P1' => ['int', 12], '@P2@P3@P4' => ['int', 19]}, HASH);
push (@testres, compare($expect, $result));

blurb("Expanding of '???' only");
if ($sqlver >= 10) {
   $expect = qr/Must declare the scalar variable ['"]\@P1\@P2(\@P3)?['"]/;
}
else {
   $expect = qr/Incorrect syntax near ['"]\@P1\@P2(\@P3)?["']/;
}
delete $X->{ErrInfo}{Messages};
$X->sql("???", {'@P1@P2@P3' => ['int', 12]});
push(@testres, ($X->{ErrInfo}{Messages}[0]{'text'} =~ $expect ? 1 : 0));

blurb("nchar/nvarchar");
$expect = [{'a' => "x'z\x{ABCD}E''F", 'b' => "\x{ABCD}EF  0xAB",
            'c' => "0xAB0xAB"}];
$result = $X->sql('SELECT a = ? + ?, b = ? + ?, c = ? + ?',
                  [['nchar', "x'z"],
                   ['nchar', "\x{ABCD}E''F"],
                   ['nchar(5)', "\x{ABCD}EF"],
                   ['nvarchar( 5)', '0xAB'],
                   ['nchar( 4 )', '0xABCDEF'],
                   ['nchar(4 )', '0xABCDEF']]);
push (@testres, compare($expect, $result));

blurb("too long nvarchar");
if ($sqlver <= 8) {
   $expect = {'b' => "21 PA\x{0179}DZIERNIKA 2004 " x 190 .
                     "21 PA\x{0179}DZIE"};
}
else {
   $expect = {'b' => "21 PA\x{0179}DZIERNIKA 2004 " x 250};
}
$result = $X->sql_one('SELECT b = upper(?)',
              [['nvarchar', "21 pa\x{017A}dziernika 2004 " x 250]],
              HASH);
push (@testres, compare($expect, $result));

blurb("too long nchar");
$expect = {'b' => "21 PA\x{0179}DZIERNIKA 2004 " x 190 .
                  "21 PA\x{0179}DZIE"};
$result = $X->sql_one('SELECT b = upper(?)',
              [['nchar', "21 pa\x{017A}dziernika 2004 " x 250]],
              HASH);
push (@testres, compare($expect, $result));


blurb ("named decimal with variations");
$X->{DecimalAsStr} = 1;
$expect = [{'d1' => '246246246246',
            'd2' => '3578',
            'd3' => '246246246246.246246',
            'n1' => '94.22',
            'n2' => '94.22'}];
$result = $X->sql('SELECT d1 = @d1 + @d1, d2 = 2*@d2, d3 = 2*@d3, n1 = @n2 + @n1, n2 = @n2 + @n2',
                  {d1 => ['decimal',          '123123123123.123123'],
                   d2 => ['decimal(10)',      '1789.44'],
                   d3 => ['decimal(18,6)',    '123123123123.123123'],
                   n1 => ['numeric(8, 2) ',   '47.11'],
                   n2 => ['numeric( 8 , 2 )', '47.11']});
push(@testres, compare($expect, $result));

blurb("named binary as binary");
$X->{BinaryAsStr} = 0;
$expect = [{'a' => "ABCDEFABCDEF", 'b' => "ABCDEF\0\0ABCDEF\0\0", 'c' => "ABCABC"}];
$result = $X->sql('SELECT a = @b1 + @b1, b = @b2 + @b2, c = @b3 + @b3',
                  {'@b1' => ['binary', 'ABCDEF'],
                   '@b2' => ['binary(8)', 'ABCDEF'],
                   '@b3' => ['binary( 3 )', 'ABCDEF']});
push (@testres, compare($expect, $result));

blurb("named char/varchar");
$expect = {'a' => "x''z 0xABCDEF", 'b' => "0xABCDEF  x''z ",
           'c' => "0xABCDEF0xABCDEF  "};
$result = $X->sql_one('SELECT a = @v1 + @v3, b = @v2 + @v1, c = @v3 + @v2',
                  {v1    => ['varchar',  "x''z "],
                   v2    => ['char(10)', '0xABCDEF'],
                   '@v3' => ['varchar( 10)', '0xABCDEF']}, HASH);
push (@testres, compare($expect, $result));

blurb("mix of named and positional parameters");
$expect = {'a' => 12, 'b' => 233, 'c' => 288};
$result = $X->sql_one('SELECT a = ?, b = @x + @y, c = ? + @x',
                      [['int', 12], ['int', 98]],
                       {'@x' => ['int', 190],
                        y    => ['int', 43]}, HASH);
push (@testres, compare($expect, $result));

$no_of_tests = 31;

if ($sqlver <= 8) {
   goto finally;
}

blurb("varchar(MAX)");
$expect = {'a' => 'HELLO DOLLY! ' x 1854,
           'b' => "21 PA\x{0179}DZIERNIKA 2004 " x 1711};
$result = $X->sql_one('SELECT a = upper(?), b = upper(?)',
              [['varchar(MAX)',  'Hello Dolly! ' x 1854],
               ['nvarchar(max)', "21 pa\x{017A}dziernika 2004 " x 1711]],
              HASH);
push (@testres, compare($expect, $result));

blurb("varbinary(MAX) as str");
$X->{BinaryAsStr} = 1;
$expect = {'a' => '0201AB60961147' x 7453,
           'b' => '47119660AB0102' x 10906};
$result = $X->sql_one('SELECT a = convert(varbinary(MAX), reverse(@b1)),
                              b = @b2 + @b2',
              {'@b1' => ['varbinary( MAX )',  '47119660AB0102' x 7453],
               '@b2' => ['varbinary(Max )', '47119660AB0102' x 5453]},
              HASH);
push (@testres, compare($expect, $result));

blurb("varbinary(MAX) as binary");
$X->{BinaryAsStr} = 0;
$expect = {'a' => '2010BA06691174' x 7453,
           'b' => '47119660AB0102' x 10906};
$result = $X->sql_one('SELECT a = convert(varbinary(MAX), reverse(@b1)),
                              b = @b2 + @b2',
              {'@b1' => ['varbinary( MaX )',  '47119660AB0102' x 7453],
               '@b2' => ['varbinary( Max)', '47119660AB0102' x 5453]},
              HASH);
push (@testres, compare($expect, $result));

blurb("XML without schema");
$expect = '<Robyn>My wife and my dead wife</Robyn>';
my $sqltext = <<'SQLEND';
SET @a.modify(N'replace value of (/Robyn/text())[1]
               with concat((/Robyn/text())[1], " and my dead wife")');
SELECT a = @a
SQLEND
$result = $X->sql_one($sqltext, {'@a' => ['xml', '<Robyn>My wife</Robyn>']},
                      SCALAR);
push (@testres, compare($expect, $result));

&blurb("XML with charset decl utf-16");
my $xml = '<?xml version="1.0" encoding="utf-16"?>' . "\n" .
          '<MMV>' . "27 pa\x{017A}dziernika 2005 " x 2000 . '</MMV>';
$expect = '<MMV>' . "27 pa\x{017A}dziernika 2005 " x 2001 . '</MMV>';
$sqltext = <<SQLEND;
SET \@a.modify(N'replace value of (/MMV/text())[1]
                with concat((/MMV/text())[1], "27 pa\x{017A}dziernika 2005 ")');
SELECT a = \@a
SQLEND
$result = $X->sql_one($sqltext, {'@a' => ['xml', $xml]}, SCALAR);
push (@testres, compare($expect, $result));

&blurb("XML with charset decl utf-8");
$xml = '<?xml version="1.0" encoding="utf-8"?>' . "\n" .
          '<MMV>' . "27 pa\x{017A}dziernika 2005 " x 2000 . '</MMV>';
$expect = '<MMV>' . "27 pa\x{017A}dziernika 2005 " x 2001 . '</MMV>';
$sqltext = <<SQLEND;
SET \@a.modify(N'replace value of (/MMV/text())[1]
                 with concat((/MMV/text())[1], "27 pa\x{017A}dziernika 2005 ")');
SELECT a = \@a
SQLEND
$result = $X->sql($sqltext, {'@a' => ['xml', $xml]}, SCALAR, SINGLEROW);
push (@testres, compare($expect, $result));

&blurb("XML with charset decl iso-8859-1");
$xml = '<?xml version="1.0" encoding="iso-8859-1"?>' . "\n" .
          '<ñandú>' . "Räksmörgås" . '</ñandú>';
$expect = '<ñandú>' . "RäksmörgåsRäksmörgås". '</ñandú>';
$sqltext = <<'SQLEND';
SET @a.modify(N'replace value of (/ñandú/text())[1]
               with concat((/ñandú/text())[1], "Räksmörgås")');
SELECT a = @a
SQLEND
$result = $X->sql_one($sqltext, {'@a' => ['xml', $xml]}, SCALAR);
push (@testres, compare($expect, $result));

$no_of_tests += 7;
goto finally if $X->{Provider} == PROVIDER_SQLOLEDB;

# Create schema for the XML with schema collection stuff
sql(<<SQLEND);
IF EXISTS (SELECT * FROM sys.xml_schema_collections WHERE name = 'OlleSC')
   DROP XML SCHEMA COLLECTION OlleSC
CREATE XML SCHEMA COLLECTION OlleSC AS '
<schema xmlns="http://www.w3.org/2001/XMLSchema">
      <element name="Olle" type="string"/>
</schema>
'
SQLEND


blurb("Testing XML with schema in parens");
$expect = '<Olle>Mors lilla Olle i skogen gick</Olle>';
$sqltext = <<'SQLEND';
SET @a.modify(N'replace value of (/Olle)[1]
               with concat((/Olle)[1], " i skogen gick")');
SELECT a = @a
SQLEND
$result = $X->sql_one($sqltext,
                      {'@a' => ['xml(OlleSC)', '<Olle>Mors lilla Olle</Olle>']},
                      SCALAR);
push (@testres, compare($expect, $result));

blurb("Testing XML with schema in parens with spaces");
$expect = '<Olle>Rosor på kinden solsken i blick</Olle>';
$sqltext = <<'SQLEND';
SET @a.modify(N'replace value of (/Olle)[1]
               with concat((/Olle)[1], " solsken i blick")');
SELECT a = @a
SQLEND
$result = $X->sql_one($sqltext,
                      {'@a' => ['XML( dbo.OlleSC )', '<Olle>Rosor på kinden</Olle>']},
                      SCALAR);
push (@testres, compare($expect, $result));

blurb("XML with schema as third param");
$expect = '<Olle>Läpparna små utav bär äro blå</Olle>';
$xml = '<?xml version="1.0" encoding="iso-8859-1"?><Olle>Läpparna små</Olle>';
$sqltext = <<'SQLEND';
SET @a.modify(N'replace value of (/Olle)[1]
               with concat((/Olle)[1], " utav bär äro blå")');
SELECT a = @a
SQLEND
$result = $X->sql_one($sqltext, {'@a' => ['xmL', $xml, 'OlleSC']}, SCALAR);
push (@testres, compare($expect, $result));

blurb("XML with schema as third param from other DB");
$X->sql("USE master");
$expect = '<Olle>Bara jag slapp att så ensam här gå</Olle>';
$sqltext = <<'SQLEND';
SET @a.modify(N'replace value of (/Olle)[1]
               with concat((/Olle)[1], " att så ensam här gå")');
SELECT a = @a
SQLEND
$result = $X->sql_one($sqltext,
            {'@a' => ['xml', '<Olle>Bara jag slapp</Olle>',
                     'tempdb..OlleSC']},
             SCALAR);
push (@testres, compare($expect, $result));
$X->sql("USE tempdb");

blurb("XML with schema in both places");
$expect = '<Olle>Mors lilla Olle med solsken i blick</Olle>';
$sqltext = <<'SQLEND';
SET @a.modify(N'replace value of (/Olle)[1]
               with concat((/Olle)[1], " med solsken i blick")');
SELECT a = @a
SQLEND
$result = $X->sql_one($sqltext,
               {'@a' => ['xml(OlleSC)', '<Olle>Mors lilla Olle</Olle>', 'OlleSC']},
               SCALAR);
push (@testres, compare($expect, $result));

$X->sql("DROP XML SCHEMA COLLECTION OlleSC");
$no_of_tests += 5;

my $clr_enabled = sql_one(<<SQLEND, Win32::SqlServer::SCALAR);
SELECT value
FROM   sys.configurations
WHERE  name = 'clr enabled'
SQLEND

goto finally if not $clr_enabled;

create_the_udts($X, 'OlleComplexInteger', 'OllePoint', 'Olle-String');

blurb("UDT with BinAsStr");
$X->{BinaryAsStr} = 1;
$expect = {p => '01800000048000000580000009',
           s => '0005000000455353494E'};
$sqltext = <<'SQLEND';
SET @p.Transpose()
SET @s = upper(@s.ToString())
SELECT p = @p, s = @s
SQLEND
$result = $X->sql_one($sqltext,
            {'@p' => ['UDT(OllePoint)', '0x01800000098000000480000005'],
             '@s' => ['UDT', '0x0005000000657373694E', '[Olle-String]']},
            HASH);
push (@testres, compare($expect, $result));

blurb("UDT with BinAsOx");
$X->{BinaryAsStr} = 'x';
$expect = {p => '0x01800000048000000580000009',
           s => '0x0005000000455353494E'};
$sqltext = <<'SQLEND';
SET @p.Transpose()
SET @s = upper(@s.ToString())
SELECT p = @p, s = @s
SQLEND
$result = $X->sql_one($sqltext,
            {'@p' => ['UDT(OllePoint)', '0x01800000098000000480000005', 'OllePoint'],
             '@s' => ['udt(dbo.[Olle-String] )', '0x0005000000657373694E']},
            HASH);
push (@testres, compare($expect, $result));

blurb("UDT with BinAsBin");
$X->{BinaryAsStr} = 0;
$expect = {p => pack('H*', '01800000048000000580000009'),
           s => pack('H*', '0005000000455353494E')};
$sqltext = <<'SQLEND';
SET @p.Transpose()
SET @s = upper(@s.ToString())
SELECT p = @p, s = @s
SQLEND
$result = $X->sql_one($sqltext,
            {'@p' => ['UDT(dbo.OllePoint)',
                       pack('H*', '01800000098000000480000005')],
             '@s' => ['UDT',
                       pack('H*', '0005000000657373694E'), '  [Olle-String]  ']},
            HASH);
push (@testres, compare($expect, $result));

blurb("UDT with BinAsStr, from other db");
$X->{BinaryAsStr} = 1;
$X->sql("USE master");
$expect = {p => '01800000048000000580000009',
           s => '0005000000455353494E'};
$sqltext = <<'SQLEND';
SET @p.Transpose()
SET @s = upper(@s.ToString())
SELECT p = @p, s = @s
SQLEND
$result = $X->sql_one($sqltext,
            {'@p' => ['UDT(tempdb.dbo.OllePoint)', '0x01800000098000000480000005'],
             '@s' => ['UDT', '0x0005000000657373694E', 'tempdb..[Olle-String]']},
            HASH);
push (@testres, compare($expect, $result));

$X->sql("USE tempdb");
delete_the_udts($X);

$no_of_tests += 4;

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
