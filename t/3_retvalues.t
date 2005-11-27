#---------------------------------------------------------------------
# $Header: /Perl/OlleDB/t/3_retvalues.t 6     05-11-26 23:47 Sommar $
#
# This test suite tests return values from sql_sp. Most of the tests
# concerns UDFs.
#
# $History: 3_retvalues.t $
# 
# *****************  Version 6  *****************
# User: Sommar       Date: 05-11-26   Time: 23:47
# Updated in $/Perl/OlleDB/t
# Renamed the module from MSSQL::OlleDB to Win32::SqlServer.
#
# *****************  Version 5  *****************
# User: Sommar       Date: 05-10-29   Time: 22:14
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 4  *****************
# User: Sommar       Date: 05-10-25   Time: 22:57
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 3  *****************
# User: Sommar       Date: 05-07-25   Time: 0:40
# Updated in $/Perl/OlleDB/t
# Added tests fpt UDT and XML.
#
# *****************  Version 2  *****************
# User: Sommar       Date: 05-06-27   Time: 22:59
# Updated in $/Perl/OlleDB/t
# Added checks for the MAX datatypes.
#
# *****************  Version 1  *****************
# User: Sommar       Date: 05-02-06   Time: 20:45
# Created in $/Perl/OlleDB/t
#---------------------------------------------------------------------

use strict;
use Win32::SqlServer qw(:DEFAULT :consts);
use File::Basename qw(dirname);

require &dirname($0) . '\testsqllogin.pl';
require '..\helpers\assemblies.pl';

use vars qw(@testres $verbose $retvalue $no_of_tests);
use constant TESTUDF => 'olledb_testudf';

sub blurb{
    push (@testres, "#------ Testing @_ ------\n");
    print "#------ Testing @_ ------\n" if $verbose;
}

sub create_udf {
    my($X, $datatype, $param, $retvalue, $prelude) = @_;
    my $testudf = TESTUDF;
    delete $X->{procs}{$testudf};
    $prelude = '' if not defined $prelude;
    $X->sql("IF object_id('$testudf') IS NOT NULL DROP FUNCTION $testudf");
    $X->sql(<<SQLEND);
    CREATE FUNCTION $testudf ($param) RETURNS $datatype AS
    BEGIN
       $prelude
       RETURN $retvalue;
    END
SQLEND
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
$X->sql("SET QUOTED_IDENTIFIER ON");
my ($sqlver) = split(/\./, $X->{SQL_version});

# Set up ErrInfo for test of return values.
$X->{errInfo}{RetStatOK}{4711}++;

$X->sql(<<'SQLEND');
IF object_id('check_ret_value') IS NOT NULL
   DROP PROCEDURE check_ret_value
IF object_id('multi_param_sp') IS NOT NULL
   DROP PROCEDURE multi_param_sp
SQLEND

$X->sql(<<'SQLEND');
CREATE PROCEDURE check_ret_value @ret int AS
   RETURN @ret
SQLEND

$X->sql(<<'SQLEND');
CREATE PROCEDURE multi_param_sp @p1 int = NULL,
                                  @p2 int = NULL OUTPUT,
                                  @p3 int = NULL,
                                  @p4 int = NULL OUTPUT,
                                  @p5 int = NULL OUTPUT AS
   SELECT @p2 = coalesce(@p1, 19) + 20
   SELECT @p4 = coalesce(@p3, 18) + 20
   SELECT @p5 = coalesce(@p5, 17) + 20
   RETURN
SQLEND

$retvalue = 233;
blurb('SP returns 0');
$X->sql_sp('check_ret_value', \$retvalue, [0]);
push(@testres, $retvalue == 0);

blurb('SP returns good non-zero');
$X->sql_sp('check_ret_value', \$retvalue, [4711]);
push(@testres, $retvalue == 4711);

blurb('SP returns bad non-zero');
eval(q!$X->sql_sp('check_ret_value', \$retvalue, [10])!);
push(@testres, ($@ =~ /returned status 10/i ? 1 : 0));

$no_of_tests = 3;

# Tests of omitting input parameters.
blurb("Input all parameters");
{ my ($p1, $p2, $p3, $p4, $p5) = (1, 2, 3, 4, 5);
  $X->sql_sp("multi_param_sp", [\$p1, \$p2, \$p3], {p4 => \$p4, p5 => \$p5});
  push(@testres, ($p2 == 21 and $p4 == 23 and $p5 == 25));
}

blurb("Only p1 and p2");
{ my ($p1, $p2) = (1, 2);
  $X->sql_sp("multi_param_sp", {p1 => \$p1, p2 => \$p2});
  push(@testres, $p2 == 21);
}

blurb("Only p3 and p4");
{ my ($p3, $p4) = (3, 4);
  $X->sql_sp("multi_param_sp", {p3 => \$p3, p4 => \$p4});
  push(@testres, $p4 == 23);
}

blurb("Only p1, p2 and p5");
{ my ($p1, $p2, $p5) = (1, undef, undef);
  $X->sql_sp("multi_param_sp", [\$p1, \$p2], {p5 => \$p5});
  push(@testres, ($p2 == 21 and $p5 == 37));
}

$no_of_tests += 4;

$X->sql(<<'SQLEND');
IF object_id('check_ret_value') IS NOT NULL
   DROP PROCEDURE check_ret_value
IF object_id('multi_param_sp') IS NOT NULL
   DROP PROCEDURE multi_param_sp
SQLEND

# For versions before SQL 2000, there is not much to test.
goto finally if ($sqlver < 8);

create_udf($X, 'bit', '', 1);
blurb('UDF bit');
$X->sql_sp(TESTUDF, \$retvalue);
push(@testres, $retvalue == 1);

create_udf($X, 'tinyint', '@param tinyint', '@param + 1');
blurb('UDF tinyint');
$X->sql_sp(TESTUDF, \$retvalue, [123]);
push(@testres, $retvalue == 124);

create_udf($X, 'smallint', '@param smallint', '-3 * @param');
blurb('UDF smallint');
$X->sql_sp(TESTUDF, \$retvalue, [123]);
push(@testres, $retvalue == -369);

create_udf($X, 'int', '@param1 int, @param2 int', '@param1 + @param2');
blurb('UDF int');
$X->sql_sp(TESTUDF, \$retvalue, {param1 => 123, param2 => -500000});
push(@testres, $retvalue == (123 - 500000));

$X->{DecimalAsStr} = 0;
create_udf($X, 'bigint', '', '123456789123456');
blurb('UDF bigint');
$X->sql_sp(TESTUDF, \$retvalue);
push(@testres, abs($retvalue - 123456789123456) < 100);

$X->{DecimalAsStr} = 1;
$retvalue = 123;
blurb('UDF bigint, decimalasstr');
$X->sql_sp(TESTUDF, \$retvalue);
push(@testres, $retvalue eq '123456789123456');

$X->{DecimalAsStr} = 0;
$retvalue = 123;
create_udf($X, 'decimal(24,6)', '@param decimal(8,6)', '@param + 123456789123456');
blurb('UDF decimal');
$X->sql_sp(TESTUDF, \$retvalue, ['0.123456']);
push(@testres, abs($retvalue - 123456789123456.123456) < 100);

$X->{DecimalAsStr} = 1;
$retvalue = 123;
blurb('UDF decimal, decimalasstr');
$X->sql_sp(TESTUDF, \$retvalue, ['0.123456']);
push(@testres, $retvalue eq '123456789123456.123456');

$retvalue = 123;
create_udf($X, 'float', '@param int', 'sqrt(@param)');
blurb('UDF float');
$X->sql_sp(TESTUDF, \$retvalue, [19]);
push(@testres, abs($retvalue - sqrt(19)) < 1E-15);

$X->{BinaryAsStr} = 0;
create_udf($X, 'binary(8)', '', '0x414243444546');
blurb('UDF binary as binary');
$X->sql_sp(TESTUDF, \$retvalue);
push(@testres, $retvalue eq "ABCDEF\x00\x00");

$X->{BinaryAsStr} = 1;
blurb('UDF binary as string');
$X->sql_sp(TESTUDF, \$retvalue);
push(@testres, $retvalue eq "4142434445460000");

$X->{BinaryAsStr} = 'x';
blurb('UDF binary as 0x');
$X->sql_sp(TESTUDF, \$retvalue);
push(@testres, $retvalue eq "0x4142434445460000");

$retvalue = 123;
create_udf($X, 'uniqueidentifier', '', "'B8581AEF-059F-4B02-934D-C15F6C9638E7'");
blurb('UDF uniqueidentifier');
$X->sql_sp(TESTUDF, \$retvalue);
push(@testres, $retvalue eq '{B8581AEF-059F-4B02-934D-C15F6C9638E7}');

create_udf($X, 'varchar(20)', "\@param varchar(20) = 'Kamel'", '@param');
blurb('UDF varchar');
$X->sql_sp(TESTUDF, \$retvalue);
push(@testres, $retvalue eq 'Kamel');

blurb('UDF varchar, empty');
$X->sql_sp(TESTUDF, \$retvalue, ['']);
push(@testres, $retvalue eq '');

blurb('UDF varchar, NULL');
$X->sql_sp(TESTUDF, \$retvalue, [undef]);
push(@testres, not defined $retvalue);

create_udf($X, 'char(20)', "\@param varchar(20) = 'Kamel'", '@param');
blurb('UDF char');
$X->sql_sp(TESTUDF, \$retvalue);
push(@testres, $retvalue eq 'Kamel' . ' ' x 15);

create_udf($X, 'nvarchar(20)', '', 'nchar(0x7623) + nchar(0x01AB) + nchar(0x2323)');
blurb('UDF nvarchar');
$X->sql_sp(TESTUDF, \$retvalue);
push(@testres, $retvalue eq "\x{7623}\x{01AB}\x{2323}");

create_udf($X, 'nchar(20)', '', 'nchar(0x7623) + nchar(0x01AB) + nchar(0x2323)');
blurb('UDF nvarchar');
$X->sql_sp(TESTUDF, \$retvalue);
push(@testres, $retvalue eq "\x{7623}\x{01AB}\x{2323}" . ' ' x 17);

$X->{DatetimeOption} = DATETIME_ISO;
create_udf($X, 'datetime', '', "'20050206 18:17:11.043'");
blurb('UDF datetime iso');
$X->sql_sp(TESTUDF, \$retvalue);
push(@testres, $retvalue eq '2005-02-06 18:17:11.043');

$X->{DatetimeOption} = DATETIME_STRFMT;
blurb('UDF datetime iso');
$X->sql_sp(TESTUDF, \$retvalue);
push(@testres, $retvalue eq '20050206 18:17:11.043');

$X->{DatetimeOption} = DATETIME_HASH;
blurb('UDF datetime iso');
$X->sql_sp(TESTUDF, \$retvalue);
push(@testres, datehash_compare($retvalue, {Year => 2005, Month => 2, Day => 6,
                                        Hour => 18, Minute => 17, Second => 11,
                                        Fraction => 43}));

create_udf($X, 'sql_variant', '@param int', 'convert(varchar, @param)');
blurb('UDF sql_variant/varchar');
undef $retvalue;
$X->sql_sp(TESTUDF, \$retvalue, [129]);
push(@testres, $retvalue eq '129');

$X->{DatetimeOption} = DATETIME_ISO;
create_udf($X, 'sql_variant', '@param int', 'convert(datetime, @param)');
blurb('UDF sql_variant/datetime');
$X->sql_sp(TESTUDF, \$retvalue, [12]);
push(@testres, $retvalue eq '1900-01-13 00:00:00.000');

$X->{BinaryAsStr} = 'x';
create_udf($X, 'sql_variant', '@param int', 'convert(varbinary(29), @param)');
blurb('UDF sql_variant/varbinary');
$X->sql_sp(TESTUDF, \$retvalue, [254]);
push(@testres, $retvalue eq '0x000000FE');

blurb('UDF sql_variant/NULL');
$X->sql_sp(TESTUDF, \$retvalue, [undef]);
push(@testres, not defined $retvalue);

$no_of_tests += 26;

if ($sqlver <= 8) {
   goto finally;
}

if ($X->{Provider} != PROVIDER_SQLNCLI) {
    goto finally;
}

blurb('UDF varchar(MAX)');
create_udf($X, 'varchar(MAX)', '@param varchar(MAX)', 'reverse(@param)');
$X->sql_sp(TESTUDF, \$retvalue, ['Palsternacksproducent' x 5000]);
push(@testres, $retvalue eq 'tnecudorpskcanretslaP' x 5000);

blurb('UDF nvarchar(MAX)');
create_udf($X, 'nvarchar(MAX)', '@param nvarchar(MAX)', 'upper(@param)');
$X->sql_sp(TESTUDF, \$retvalue, ["21 pa\x{017A}dziernika 2004 " x 5000]);
push(@testres, $retvalue eq "21 PA\x{0179}DZIERNIKA 2004 " x 5000);

$X->{BinaryAsStr} = 1;
blurb('UDF varbinary(MAX) as str');
create_udf($X, 'varbinary(MAX)', '@param varbinary(MAX)',
                'convert(varbinary(MAX), reverse(@param))');
$X->sql_sp(TESTUDF, \$retvalue, ['0x' . '47119660AB0102' x 5000]);
push(@testres, $retvalue eq '0201AB60961147' x 5000);

$X->{BinaryAsStr} = 'x';
blurb('UDF varbinary(MAX) 0x');
$X->sql_sp(TESTUDF, \$retvalue, ['47119660AB0102' x 5000]);
push(@testres, $retvalue eq '0x' . '0201AB60961147' x 5000);

$X->{BinaryAsStr} = 0;
blurb('UDF varbinary(MAX) as bin');
$X->sql_sp(TESTUDF, \$retvalue, ['47119660AB0102' x 5000]);
push(@testres, $retvalue eq '2010BA06691174' x 5000);

blurb('XML');
create_udf($X, 'xml', '@param xml', '@param',
           q!SET @param.modify('replace value of (/TEST/text())[1]
                                with concat((/TEST/text())[1], " extra text")')!);
$X->sql_sp(TESTUDF, \$retvalue, ['<TEST>regular text</TEST>']);
push(@testres, $retvalue eq '<TEST>regular text extra text</TEST>');

blurb('Long XML');
my $longoctober = "21 pa\x{017A}dziernika 2004 " x 2000;
create_udf($X, 'xml', '@param nvarchar(MAX)', '@xml',
           q!DECLARE @xml xml; SET @xml = (SELECT @param AS x FOR XML RAW)!);
$X->sql_sp(TESTUDF, \$retvalue, [$longoctober]);
push(@testres,
     ($retvalue =~ m!^<row\s+x\s*=\s*\"$longoctober\"\s*/\s*>$! ? 1 : 0));

blurb('XML(OlleSC)');
sql(<<SQLEND);
IF EXISTS (SELECT * FROM sys.xml_schema_collections WHERE name = 'OlleSC')
   DROP XML SCHEMA COLLECTION OlleSC
SQLEND
sql(<<SQLEND);
CREATE XML SCHEMA COLLECTION OlleSC AS '
<schema xmlns="http://www.w3.org/2001/XMLSchema">
      <element name="root" type="string"/>
</schema>
'
SQLEND
create_udf($X, 'xml(OlleSC)', '@param xml(OlleSC)', '@param',
           q!SET @param.modify('replace value of (/root)[1]
                                with concat((/root)[1], " added text")')!);
$X->sql_sp(TESTUDF, \$retvalue, ['<root>initial text</root>']);
push(@testres, $retvalue eq '<root>initial text added text</root>');

sql(<<SQLEND);
DROP FUNCTION olledb_testudf
IF EXISTS (SELECT * FROM sys.xml_schema_collections WHERE name = 'OlleSC')
   DROP XML SCHEMA COLLECTION OlleSC
SQLEND

$no_of_tests += 8;


my $clr_enabled = sql_one(<<SQLEND, Win32::SqlServer::SCALAR);
SELECT value
FROM   sys.configurations
WHERE  name = 'clr enabled'
SQLEND

goto finally if not $clr_enabled;

create_the_udts($X, 'OlleComplexInteger', 'Olle.Point', 'OlleString');
$X->{BinaryAsStr} = 'x';
blurb('UDT1, bin0x');
create_udf($X, '[Olle.Point]', '@p [Olle.Point]', '@p', 'SET @p.Transpose()');
$X->sql_sp(TESTUDF, \$retvalue, ['0x01800000098000000480000005']);
push(@testres, $retvalue eq      '0x01800000048000000580000009');

$X->{BinaryAsStr} = 0;
blurb('UDT3, binary as binary');
create_udf($X, 'OlleString', '@s OlleString', 'upper(@s.ToString())');
$X->sql_sp(TESTUDF, \$retvalue, [pack('H*', '0005000000657373694E')]);
push(@testres, $retvalue eq pack('H*', '0005000000455353494E'));


$no_of_tests += 2;


finally:

if ($sqlver >= 8) {
   my $testudf = TESTUDF;
   $X->sql(<<SQLEND);
   IF EXISTS (SELECT * FROM sysobjects WHERE name = '$testudf')
      DROP FUNCTION $testudf
SQLEND
}

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
