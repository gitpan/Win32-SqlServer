#---------------------------------------------------------------------
# $Header: /Perl/OlleDB/t/2_datatypes.t 17    05-11-26 23:47 Sommar $
#
# This test script tests using sql_sp and sql_insert in all possible
# ways and with testing use of all datatypes.
#
# $History: 2_datatypes.t $
# 
# *****************  Version 17  *****************
# User: Sommar       Date: 05-11-26   Time: 23:47
# Updated in $/Perl/OlleDB/t
# Renamed the module from MSSQL::OlleDB to Win32::SqlServer.
#
# *****************  Version 16  *****************
# User: Sommar       Date: 05-11-06   Time: 20:49
# Updated in $/Perl/OlleDB/t
# Added test for datetime format YYYY-MM-DDZ.
#
# *****************  Version 15  *****************
# User: Sommar       Date: 05-10-23   Time: 23:12
# Updated in $/Perl/OlleDB/t
# Added more tests for XML.
#
# *****************  Version 14  *****************
# User: Sommar       Date: 05-08-07   Time: 0:16
# Updated in $/Perl/OlleDB/t
# Modified the Unicode test to also include Unicode in parameter names.
#
# *****************  Version 13  *****************
# User: Sommar       Date: 05-07-25   Time: 0:39
# Updated in $/Perl/OlleDB/t
# Added clean-up code to leave nothing around.
#
# *****************  Version 12  *****************
# User: Sommar       Date: 05-07-20   Time: 22:42
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 11  *****************
# User: Sommar       Date: 05-07-18   Time: 1:00
# Updated in $/Perl/OlleDB/t
# Tests for untyped XML as well.
#
# *****************  Version 10  *****************
# User: Sommar       Date: 05-07-17   Time: 23:11
# Updated in $/Perl/OlleDB/t
# Tests for UDT added.
#
# *****************  Version 9  *****************
# User: Sommar       Date: 05-06-25   Time: 23:01
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 8  *****************
# User: Sommar       Date: 05-02-06   Time: 20:45
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 7  *****************
# User: Sommar       Date: 05-01-30   Time: 21:56
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 6  *****************
# User: Sommar       Date: 05-01-24   Time: 23:09
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 5  *****************
# User: Sommar       Date: 05-01-24   Time: 0:41
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 4  *****************
# User: Sommar       Date: 05-01-19   Time: 23:07
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 3  *****************
# User: Sommar       Date: 05-01-10   Time: 23:02
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 2  *****************
# User: Sommar       Date: 05-01-10   Time: 20:55
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 1  *****************
# User: Sommar       Date: 05-01-06   Time: 22:59
# Created in $/Perl/OlleDB/t
#
# *****************  Version 3  *****************
# User: Sommar       Date: 00-07-24   Time: 22:10
# Updated in $/Perl/MSSQL/Sqllib/t
# Changed nullif argument for bincol due to bug(?) in SQL 2000 Beta 2.
#
# *****************  Version 2  *****************
# User: Sommar       Date: 00-05-08   Time: 22:23
# Updated in $/Perl/MSSQL/Sqllib/t
# Enhanced test for text and image to use really big stuff.
#
# *****************  Version 1  *****************
# User: Sommar       Date: 99-01-30   Time: 16:36
# Created in $/Perl/MSSQL/sqllib/t
#---------------------------------------------------------------------

use strict;
use IO::File;
use English;

use vars qw($sqlver @tblcols $no_of_tests @testres %tbl
            %expectpar %expectcol %expectfile %test %filetest %comment);

use constant TESTFILE => "datatypes.log";

sub blurb{
    push(@testres, "#------ Testing @_ ------");
    print "#------ Testing @_ ------\n";
}

use Win32::SqlServer qw(:DEFAULT :consts);
use Filehandle;
use File::Basename qw(dirname);

require &dirname($0) . '\testsqllogin.pl';
require '..\helpers\assemblies.pl';

sub clear_test_data {
   @tblcols = %tbl = %expectpar = %expectcol = %expectfile =
   %test = %filetest = %comment = ();
}

sub drop_test_objects {
    my ($type) = @_;
    sql("IF object_id('$type') IS NOT NULL DROP TABLE $type");
    sql("IF object_id('${type}_sp') IS NOT NULL DROP PROCEDURE ${type}_sp");
}

sub create_integer {
   drop_test_objects('integer');

   sql(<<SQLEND);
      CREATE TABLE integer (intcol      int       NULL,
                           smallintcol  smallint  NULL,
                           tinyintcol   tinyint   NULL,
                           floatcol     float     NULL,
                           realcol      real      NULL,
                           bitcol       bit       NOT NULL)
SQLEND

   @tblcols = qw(intcol smallintcol tinyintcol floatcol realcol bitcol);

   sql(<<SQLEND);
      CREATE TRIGGER integer_tri ON integer FOR INSERT AS
      UPDATE integer
      SET    intcol      = intcol - 4711,
             smallintcol = smallintcol - 4711,
             tinyintcol  = tinyintcol - 47,
             floatcol    = floatcol - 4711,
             realcol     = realcol  - 4711,
             bitcol      = 1 - bitcol
SQLEND

   sql(<<'SQLEND');
   CREATE PROCEDURE integer_sp
                    @intcol       int           OUTPUT,
                    @smallintcol  smallint      OUTPUT,
                    @tinyintcol   tinyint       OUTPUT,
                    @floatcol     float         OUTPUT,
                    @realcol      real          OUTPUT,
                    @bitcol       bit           OUTPUT AS

   DELETE integer

   INSERT integer (intcol, smallintcol, tinyintcol, floatcol, realcol, bitcol)
      VALUES (@intcol, @smallintcol, @tinyintcol, @floatcol, @realcol,
              isnull(@bitcol, 0))

   SELECT @intcol       = -2 * @intcol,
          @smallintcol  = -2 * @smallintcol,
          @tinyintcol   =  2 * @tinyintcol,
          @floatcol     = -2 * @floatcol,
          @realcol      = -2 * @realcol,
          @bitcol       =  1 - @bitcol

   SELECT intcol, smallintcol, tinyintcol, floatcol, realcol, bitcol
   FROM   integer
SQLEND
}

sub create_character {
   drop_test_objects('character');

   sql(<<SQLEND);
      CREATE TABLE character(charcol      char(20)      NULL,
                             varcharcol   varchar(20)   NULL,
                             varcharcol2  varchar(20)   NOT NULL,
                             textcol      text          NULL);
SQLEND

   @tblcols = qw(charcol varcharcol varcharcol2 textcol);

   sql(<<SQLEND);
      CREATE TRIGGER character_tri ON character FOR INSERT AS
      UPDATE character
      SET    charcol     = reverse(charcol),
             varcharcol  = reverse(varcharcol),
             varcharcol2 = reverse(varcharcol2)
SQLEND

   sql(<<'SQLEND');
   CREATE PROCEDURE character_sp
                    @charcol     char(20)    OUTPUT,
                    @varcharcol  varchar(20) OUTPUT,
                    @varcharcol2 varchar(20) OUTPUT,
                    @textcol     text  AS

   DELETE character

   INSERT character(charcol, varcharcol, varcharcol2, textcol)
      VALUES (@charcol, @varcharcol, @varcharcol2, @textcol)

   SELECT @charcol     = upper(@charcol),
          @varcharcol  = upper(@varcharcol),
          @varcharcol2 = upper(@varcharcol2)

   SELECT charcol, varcharcol, varcharcol2, textcol
   FROM   character
SQLEND
}

sub create_binary {
   drop_test_objects('binary');

   sql(<<SQLEND);
      CREATE TABLE binary(bincol      binary(20)    NULL,
                          varbincol   varbinary(20) NULL,
                          tstamp      timestamp     NOT NULL,
                          imagecol    image         NULL);
SQLEND

   @tblcols = qw(bincol varbincol tstamp imagecol);

   sql(<<SQLEND);
      CREATE TRIGGER binary_tri ON binary FOR INSERT AS
      UPDATE binary
      SET    bincol     = convert(binary(20), reverse(bincol)),
             varbincol  = convert(varbinary(20), reverse(varbincol))
SQLEND

   sql(<<'SQLEND');
   CREATE PROCEDURE binary_sp
                    @bincol     binary(20)    OUTPUT,
                    @varbincol  varbinary(20) OUTPUT,
                    @tstamp     timestamp     OUTPUT,
                    @imagecol   image  AS

   DELETE binary

   INSERT binary(bincol, varbincol, imagecol)
      VALUES (@bincol, @varbincol, @imagecol)

   SELECT @bincol     = substring(@bincol, 1, 4) + @bincol,
          @varbincol  = @varbincol + @varbincol,
          @tstamp     = substring(@tstamp, 5, 4) + substring(@tstamp, 1, 4)

   SELECT bincol, varbincol, tstamp = @tstamp, imagecol
   FROM   binary
SQLEND
}


sub create_decimal {
   drop_test_objects('decimal');

   sql(<<SQLEND);
      CREATE TABLE decimal(deccol       decimal(24,6) NULL,
                           numcol       numeric(12,2) NULL,
                           moneycol     money         NULL,
                           dimecol      smallmoney    NULL)
SQLEND

   @tblcols = qw(deccol numcol moneycol dimecol);

   sql(<<SQLEND);
      CREATE TRIGGER decimal_tri ON decimal FOR INSERT AS
      UPDATE decimal
      SET    deccol      = deccol   - 12345678,
             numcol      = numcol   - 12345678,
             moneycol    = moneycol - 12345678,
             dimecol     = dimecol  - 123456
SQLEND

   sql(<<'SQLEND');
   CREATE PROCEDURE decimal_sp
                    @deccol       decimal(24,6) OUTPUT,
                    @numcol       numeric(12,2) OUTPUT,
                    @moneycol     money         OUTPUT,
                    @dimecol      smallmoney    OUTPUT AS

   DELETE decimal

   INSERT decimal(deccol, numcol, moneycol, dimecol)
      VALUES (@deccol, @numcol, @moneycol, @dimecol)

   SELECT @deccol   = -2 * @deccol,
          @numcol   = -1 * @numcol / 2,
          @moneycol = -2 * @moneycol,
          @dimecol  = -1 * @dimecol / 2

   SELECT deccol, numcol, moneycol, dimecol
   FROM   decimal
SQLEND
}

sub create_datetime {
   drop_test_objects('datetime');

   sql(<<SQLEND);
      CREATE TABLE datetime(datetimecol   datetime      NULL,
                            smalldatecol  smalldatetime NULL)
SQLEND

   @tblcols = qw(datetimecol smalldatecol);

   sql(<<SQLEND);
      CREATE TRIGGER datetime_tri ON datetime FOR INSERT AS
      UPDATE datetime
      SET    datetimecol  = dateadd(DAY, 17, datetimecol),
             smalldatecol = dateadd(MONTH, 3, smalldatecol)
SQLEND

   sql(<<'SQLEND');
   CREATE PROCEDURE datetime_sp
                    @datetimecol  datetime      OUTPUT,
                    @smalldatecol smalldatetime OUTPUT AS

   DELETE datetime

   INSERT datetime(datetimecol, smalldatecol)
      VALUES (@datetimecol, @smalldatecol)

   SELECT @datetimecol   = dateadd(HOUR,    4, @datetimecol),
          @smalldatecol  = dateadd(MINUTE, 14, @smalldatecol)

   SELECT datetimecol, smalldatecol
   FROM   datetime
SQLEND
}

sub create_guid {
   drop_test_objects('guid');

   sql(<<SQLEND);
      CREATE TABLE guid(guidcol    uniqueidentifier NULL,
                        nullbitcol bit              NULL)
SQLEND

   @tblcols = qw(guidcol nullbitcol);

   sql(<<SQLEND);
      CREATE TRIGGER guid_tri ON guid FOR INSERT AS
      UPDATE guid
      SET    guidcol    = convert(uniqueidentifier,
                            replace(convert(char(36), guidcol), 'F', '0')),
             nullbitcol = 1 - nullbitcol
SQLEND

   sql(<<'SQLEND');
   CREATE PROCEDURE guid_sp
                    @guidcol     uniqueidentifier OUTPUT,
                    @nullbitcol  bit OUTPUT AS

   DELETE guid

   INSERT guid(guidcol, nullbitcol)
      VALUES (@guidcol, @nullbitcol)

   SELECT @guidcol    = convert(uniqueidentifier,
                            replace(convert(char(36), @guidcol), 'F', 'A')),
          @nullbitcol = 1 - @nullbitcol

   SELECT guidcol, nullbitcol
   FROM   guid
SQLEND
}

sub create_unicode {
   drop_test_objects('unicode');

   sql(<<SQLEND);
      CREATE TABLE unicode(ncharcol             nchar(20)    NULL,
                           \x{0144}varcharcol   nvarchar(20) NULL,
                           nchärcöl2            nchar(20)    NOT NULL,
                           ntextcol             ntext        NULL);
SQLEND

   @tblcols = ("ncharcol", "\x{0144}varcharcol", "nchärcöl2", "ntextcol");

   sql(<<SQLEND);
      CREATE TRIGGER unicode_tri ON unicode FOR INSERT AS
      UPDATE unicode
      SET    ncharcol     = reverse(ncharcol),
             \x{0144}varcharcol  = reverse(\x{0144}varcharcol),
             nchärcöl2    = reverse(nchärcöl2)
SQLEND

   sql(<<SQLEND);
   CREATE PROCEDURE unicode_sp
                    \@ncharcol     nchar(20)    OUTPUT,
                    \@\x{0144}varcharcol  nvarchar(20) OUTPUT,
                    \@nchärcöl2    nchar(20)    OUTPUT,
                    \@ntextcol     ntext  AS

   DELETE unicode

   INSERT unicode(ncharcol, \x{0144}varcharcol, nchärcöl2, ntextcol)
      VALUES (\@ncharcol, \@\x{0144}varcharcol, \@nchärcöl2, \@ntextcol)

   SELECT \@ncharcol     = upper(\@ncharcol),
          \@\x{0144}varcharcol  = upper(\@\x{0144}varcharcol),
          \@nchärcöl2    = upper(\@nchärcöl2)

   SELECT ncharcol, \x{0144}varcharcol, nchärcöl2, ntextcol
   FROM   unicode
SQLEND
}

sub create_bigint {
   drop_test_objects('bigint');

   sql(<<SQLEND);
      CREATE TABLE bigint(bigintcol bigint NULL)
SQLEND

   @tblcols = qw(bigintcol);

   sql(<<SQLEND);
      CREATE TRIGGER bigint_tri ON bigint FOR INSERT AS
      UPDATE bigint
      SET    bigintcol = bigintcol   - 12345678
SQLEND

   sql(<<'SQLEND');
   CREATE PROCEDURE bigint_sp @bigintcol  bigint OUTPUT AS

   DELETE bigint

   INSERT bigint(bigintcol)
      VALUES (@bigintcol)

   SELECT @bigintcol = -2 * @bigintcol

   SELECT bigintcol
   FROM   bigint
SQLEND
}

sub create_sql_variant {
# sql_variant is a bit different from the rest...
   drop_test_objects('sql_variant');

   sql(<<SQLEND);
      CREATE TABLE sql_variant(varcol   sql_variant  NULL,
                               intype   sysname      NULL,
                               outtype  sysname      NOT NULL)
SQLEND

   @tblcols  = qw(varcol outtype intype);

   sql(<<'SQLEND');
      CREATE TRIGGER sql_variant_tri ON sql_variant FOR INSERT AS
      DECLARE @var      sql_variant,
              @outtype  sysname
      UPDATE sql_variant
      SET    intype   = convert(sysname,
                                sql_variant_property(varcol, 'Basetype')),
             @outtype = outtype,
             @var     = varcol

      IF @outtype = 'bit'
         UPDATE sql_variant SET varcol = convert(bit, @var)
      ELSE IF @outtype = 'tinyint'
         UPDATE sql_variant SET varcol = convert(tinyint, @var) -
                                         convert(tinyint, 50)
      ELSE IF @outtype = 'smallint'
         UPDATE sql_variant SET varcol = convert(smallint, @var) -
                                         convert(smallint, 50)
      ELSE IF @outtype = 'int'
         UPDATE sql_variant SET varcol = convert(int, @var) -
                                         convert(int, 50)
      ELSE IF @outtype = 'bigint'
         UPDATE sql_variant SET varcol = convert(bigint, @var) -
                                         convert(bigint, 12345678)
      ELSE IF @outtype = 'real'
         UPDATE sql_variant SET varcol = convert(real, @var)  -
                                         convert(real, 50)
      ELSE IF @outtype = 'float'
         UPDATE sql_variant SET varcol = convert(float, @var) -
                                         convert(float, 50)
      ELSE IF @outtype = 'decimal'
         UPDATE sql_variant SET varcol = convert(decimal(24,6), @var) -
                                         convert(decimal(24,6), 12345678)
      ELSE IF @outtype = 'numeric'
         UPDATE sql_variant SET varcol = convert(numeric(12,2), @var) -
                                         convert(numeric(12,2), 12345678)
      ELSE IF @outtype = 'money'
         UPDATE sql_variant SET varcol = convert(money, @var) -
                                         convert(money, 12345678)
      ELSE IF @outtype = 'smallmoney'
         UPDATE sql_variant SET varcol = convert(smallmoney, @var) -
                                         convert(smallmoney, 12345)
      ELSE IF @outtype = 'datetime'
         UPDATE sql_variant SET varcol = dateadd(DAY, -50,
                                            convert(datetime, @var))
      ELSE IF @outtype = 'smalldatetime'
         UPDATE sql_variant SET varcol = dateadd(DAY, -50,
                                            convert(smalldatetime, @var))
      ELSE IF @outtype = 'char'
         UPDATE sql_variant SET varcol = reverse(convert(char(20), @var))
      ELSE IF @outtype = 'varchar'
         UPDATE sql_variant SET varcol = reverse(convert(varchar(20), @var))
      ELSE IF @outtype = 'nchar'
         UPDATE sql_variant SET varcol = reverse(convert(nchar(20), @var))
      ELSE IF @outtype = 'nvarchar'
         UPDATE sql_variant SET varcol = reverse(convert(nvarchar(20), @var))
      ELSE IF @outtype = 'binary'
         UPDATE sql_variant SET varcol = convert(binary(20), @var)
      ELSE IF @outtype = 'varbinary'
         UPDATE sql_variant SET varcol = convert(varbinary(20), @var)
      ELSE IF @outtype = 'uniqueidentifier'
         UPDATE sql_variant SET varcol = convert(uniqueidentifier, @var)
      ELSE
         UPDATE sql_variant SET varcol = NULL
SQLEND

   sql(<<'SQLEND');
   CREATE PROCEDURE sql_variant_sp
                    @varcol       sql_variant     OUTPUT,
                    @outtype      sysname,
                    @intype       sysname  = NULL OUTPUT AS

   DELETE sql_variant

   INSERT sql_variant(varcol, outtype)
      VALUES (@varcol, @outtype)

   SELECT @intype = convert(sysname, sql_variant_property(@varcol, 'Basetype'))

   IF @outtype = 'bit'
      SELECT @varcol = convert(bit, @varcol)
   ELSE IF @outtype = 'tinyint'
      SELECT @varcol = convert(tinyint, 2) * convert(tinyint, @varcol)
   ELSE IF @outtype = 'smallint'
      SELECT @varcol = convert(smallint, -2) * convert(smallint, @varcol)
   ELSE IF @outtype = 'int'
      SELECT @varcol = convert(int, -2) * convert(int, @varcol)
   ELSE IF @outtype = 'bigint'
      SELECT @varcol = convert(bigint, -2) * convert(bigint, @varcol)
   ELSE IF @outtype = 'real'
      SELECT @varcol = convert(real, -2) * convert(real, @varcol)
   ELSE IF @outtype = 'float'
      SELECT @varcol = convert(float, -2) * convert(float, @varcol)
   ELSE IF @outtype = 'decimal'
      SELECT @varcol = convert(decimal(5,0), -2) * convert(decimal(24,6), @varcol)
   ELSE IF @outtype = 'numeric'
      SELECT @varcol = convert(numeric(5,0), -2) * convert(numeric(12,2), @varcol)
   ELSE IF @outtype = 'money'
      SELECT @varcol = convert(money, -2) * convert(money, @varcol)
   ELSE IF @outtype = 'smallmoney'
      SELECT @varcol = convert(smallmoney, -2) * convert(smallmoney, @varcol)
   ELSE IF @outtype = 'datetime'
      SELECT @varcol = dateadd(HOUR, 10, convert(datetime, @varcol))
   ELSE IF @outtype = 'smalldatetime'
      SELECT @varcol = dateadd(HOUR, 10, convert(smalldatetime, @varcol))
   ELSE IF @outtype = 'char'
      SELECT @varcol = upper(convert(char(20), @varcol))
   ELSE IF @outtype = 'varchar'
      SELECT @varcol = upper(convert(varchar(20), @varcol))
   ELSE IF @outtype = 'nchar'
      SELECT @varcol = upper(convert(nchar(20), @varcol))
   ELSE IF @outtype = 'nvarchar'
      SELECT @varcol = upper(convert(nvarchar(20), @varcol))
   ELSE IF @outtype = 'binary'
      SELECT @varcol = convert(binary(20), @varcol)
   ELSE IF @outtype = 'varbinary'
      SELECT @varcol = convert(varbinary(20), @varcol)
   ELSE IF @outtype = 'uniqueidentifier'
      SELECT @varcol = convert(uniqueidentifier, @varcol)
   ELSE
      SELECT @varcol = NULL

   SELECT varcol, intype, outtype = @outtype
   FROM   sql_variant
SQLEND
}

sub create_varcharmax {
   drop_test_objects('varcharmax');

   sql(<<SQLEND);
      CREATE TABLE varcharmax(varcharcol   varchar(MAX)   NULL,
                              nvarcharcol  nvarchar(MAX)  NOT NULL)
SQLEND

   @tblcols = qw(varcharcol nvarcharcol);

   sql(<<SQLEND);
      CREATE TRIGGER varcharmax_tri ON varcharmax FOR INSERT AS
      UPDATE varcharmax
      SET    varcharcol   = reverse(varcharcol),
             nvarcharcol  = reverse(nvarcharcol)
SQLEND

   sql(<<'SQLEND');
   CREATE PROCEDURE varcharmax_sp
                    @varcharcol  varchar(MAX)   OUTPUT,
                    @nvarcharcol nvarchar(MAX)  OUTPUT AS

   DELETE varcharmax

   INSERT varcharmax(varcharcol, nvarcharcol)
      VALUES (@varcharcol, @nvarcharcol)

   SELECT @varcharcol  = upper(@varcharcol),
          @nvarcharcol = upper(@nvarcharcol)

   SELECT varcharcol, nvarcharcol
   FROM   varcharmax
SQLEND
}

sub create_varbinmax {
   drop_test_objects('varbinmax');

   sql(<<SQLEND);
      CREATE TABLE varbinmax(varbincol  varbinary(MAX) NULL);
SQLEND

   @tblcols = qw(varbincol);

   sql(<<SQLEND);
      CREATE TRIGGER varbinmax_tri ON varbinmax FOR INSERT AS
      UPDATE varbinmax
      SET    varbincol  = convert(varbinary(MAX), reverse(varbincol))
SQLEND

   sql(<<'SQLEND');
   CREATE PROCEDURE varbinmax_sp
                    @varbincol   varbinary(MAX) OUTPUT AS

   DELETE varbinmax

   INSERT varbinmax(varbincol)
      VALUES (@varbincol)

   SELECT @varbincol = @varbincol + @varbincol

   SELECT varbincol
   FROM   varbinmax
SQLEND
}

sub create_UDT1 {
    my($X, $output) = @_;

    drop_test_objects('UDT1');

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

    create_the_udts($X, 'OlleComplexInteger', 'OllePoint', 'OlleString');

    sql(<<SQLEND);
       CREATE TABLE UDT1 (cmplxcol  OlleComplexInteger NULL,
                          pointcol  OllePoint          NULL,
                          stringcol OlleString         NULL,
                          xmlcol    xml(OlleSC)        NULL)
SQLEND

    @tblcols = qw(cmplxcol pointcol stringcol xmlcol);

    sql(<<SQLEND);
       CREATE TRIGGER UDT1_tri ON UDT1 FOR INSERT AS
       UPDATE UDT1
       SET    cmplxcol  = '(' + str(cmplxcol.Imaginary) + ',' +
                          str(cmplxcol.Real) + 'i)',
              pointcol  = ltrim(str(2*pointcol.X)) + ':' +
                          ltrim(str(2*pointcol.Y)) + ':' +
                          ltrim(str(2*pointcol.Z)),
              stringcol = reverse(stringcol.ToString())

       UPDATE UDT1
       SET    xmlcol.modify('replace value of (/root)[1]
                             with concat((/root)[1], " trigger text")')
       WHERE  xmlcol IS NOT NULL
SQLEND

   my $spcode = <<'SQLEND';
       CREATE PROCEDURE UDT1_sp @cmplxcol  OlleComplexInteger OUTPUT,
                                @pointcol  OllePoint          OUTPUT,
                                @stringcol OlleString         OUTPUT,
                                @xmlcol    xml(OlleSC)        OUTPUT AS

       DELETE UDT1

       INSERT UDT1 (cmplxcol, pointcol, stringcol, xmlcol)
          VALUES (@cmplxcol, @pointcol, @stringcol, @xmlcol)

       IF @cmplxcol IS NOT NULL
       BEGIN
          SET @cmplxcol.Real      = 2 * @cmplxcol.Real
          SET @cmplxcol.Imaginary = 2 * @cmplxcol.Imaginary
       END

       IF @pointcol IS NOT NULL
          SET @pointcol.Transpose()

       SELECT @stringcol = UPPER(@stringcol.ToString())

       IF @xmlcol IS NOT NULL
          SET @xmlcol.modify('replace value of (/root)[1]
                             with concat((/root)[1], " procedure text")')

       SELECT cmplxcol, pointcol, stringcol, xmlcol
       FROM   UDT1
SQLEND

    if (not $output) {
       $spcode =~ s/\bOUTPUT\b//g;
    }

    sql($spcode);
}

sub create_UDT2 {
    my($X, $output) = @_;

   drop_test_objects('UDT2');

   sql(<<SQLEND);
       CREATE TABLE UDT2 (cmplxcol  OlleComplexInteger NULL,
                          intcol    int                NULL,
                          stringcol OlleString         NULL)
SQLEND

    @tblcols = qw(cmplxcol intcol stringcol);

    sql(<<SQLEND);
       CREATE TRIGGER UDT2_tri ON UDT2 FOR INSERT AS
       UPDATE UDT2
       SET    cmplxcol  = '(' + str(cmplxcol.Imaginary) + ',' +
                          str(cmplxcol.Real) + 'i)',
              intcol    = 2*intcol,
              stringcol = reverse(stringcol.ToString())
SQLEND

   my $spcode = <<'SQLEND';
       CREATE PROCEDURE UDT2_sp @cmplxcol  OlleComplexInteger OUTPUT,
                                @intcol    int                OUTPUT,
                                @stringcol OlleString         OUTPUT AS

       DELETE UDT2

       INSERT UDT2 (cmplxcol, intcol, stringcol)
          VALUES (@cmplxcol, @intcol, @stringcol)

       IF @cmplxcol IS NOT NULL
       BEGIN
          SET @cmplxcol.Real      = 2 * @cmplxcol.Real
          SET @cmplxcol.Imaginary = 2 * @cmplxcol.Imaginary
       END

       SELECT @intcol    = @intcol + 91,
              @stringcol = UPPER(@stringcol.ToString())

       SELECT cmplxcol, intcol, stringcol
       FROM   UDT2
SQLEND

    if (not $output) {
       $spcode =~ s/\bOUTPUT\b//g;
    }

    sql($spcode);
}

sub create_UDT3 {
    my($X, $output) = @_;

    drop_test_objects('UDT3');

    sql(<<SQLEND);
       CREATE TABLE UDT3 (xmlcol    xml        NULL,
                          pointcol  OllePoint  NULL,
                          nollcol   float      NULL)
SQLEND

    @tblcols = qw(xmlcol pointcol nollcol);

    sql(<<SQLEND);
       CREATE TRIGGER UDT3_tri ON UDT3 FOR INSERT AS
       UPDATE UDT3
       SET    pointcol  = ltrim(str(2*pointcol.X)) + ':' +
                          ltrim(str(2*pointcol.Y)) + ':' +
                          ltrim(str(2*pointcol.Z)),
              nollcol   = nollcol + 19

       UPDATE UDT3
       SET    xmlcol.modify('replace value of (/TEST/text())[1]
                             with concat((/TEST/text())[1], " trigger text")')
       WHERE  xmlcol IS NOT NULL
SQLEND

   my $spcode = <<'SQLEND';
       CREATE PROCEDURE UDT3_sp @xmlcol    xml       OUTPUT,
                                @pointcol  OllePoint OUTPUT,
                                @nollcol   float     OUTPUT AS

       DELETE UDT3

       INSERT UDT3 (xmlcol, pointcol, nollcol)
          VALUES (@xmlcol, @pointcol, @nollcol)

       IF @xmlcol IS NOT NULL
          SET @xmlcol.modify('replace value of (/TEST/text())[1]
                             with concat((/TEST/text())[1], " procedure text")')

       IF @pointcol IS NOT NULL
          SET @pointcol.Transpose()

       SELECT @nollcol  = @nollcol - 9

       SELECT xmlcol, pointcol, nollcol
       FROM   UDT3
SQLEND

    if (not $output) {
       $spcode =~ s/\bOUTPUT\b//g;
    }

    sql($spcode);
}

sub create_xmltest {
    my($X, $output) = @_;

    drop_test_objects('xmltest');

    sql(<<SQLEND);
    IF EXISTS (SELECT * FROM sys.xml_schema_collections WHERE name = 'Olles SC')
            DROP XML SCHEMA COLLECTION [Olles SC]
SQLEND

     sql(<<SQLEND);
CREATE XML SCHEMA COLLECTION [Olles SC] AS '
<schema xmlns="http://www.w3.org/2001/XMLSchema">
      <element name="TÄST" type="string"/>
</schema>
'
SQLEND


    sql(<<SQLEND);
       CREATE TABLE xmltest (xmlcol      xml             NULL,
                             xmlsccol    xml([Olles SC]) NULL,
                             nvarcol     nvarchar(MAX)   NULL,
                             nvarsccol   nvarchar(MAX)   NULL)
SQLEND

    @tblcols = qw(xmlcol xmlsccol nvarcol nvarsccol);

    sql(<<SQLEND);
       CREATE TRIGGER xmltest_tri ON xmltest FOR INSERT AS
       UPDATE xmltest
       SET    xmlcol    = (SELECT nvarcol FROM xmltest FOR XML AUTO),
              xmlsccol  = (SELECT 1 AS Tag, NULL as Parent,
                                  nvarsccol AS [TÄST!1]
                           FROM   xmltest
                           FOR    XML EXPLICIT),
              nvarcol   = nullif(convert(nvarchar(MAX), xmlcol), ''),
              nvarsccol = xmlsccol.value(N'/TÄST[1]', 'nvarchar(MAX)')
SQLEND

   my $spcode = <<'SQLEND';
       CREATE PROCEDURE xmltest_sp @xmlcol    xml             OUTPUT,
                                   @xmlsccol  xml([Olles SC]) OUTPUT,
                                   @nvarcol   nvarchar(MAX)   OUTPUT,
                                   @nvarsccol nvarchar(MAX)   OUTPUT AS

       DECLARE @tmp   nvarchar(MAX),
               @tmpsc nvarchar(MAX)

       DELETE xmltest

       INSERT xmltest (xmlcol, xmlsccol, nvarcol, nvarsccol)
          VALUES (@xmlcol, @xmlsccol, @nvarcol, @nvarsccol)

       SELECT @tmp = @nvarcol, @tmpsc = @nvarsccol

       SELECT @nvarcol = @xmlcol.value(N'/*[1]', 'nvarchar(MAX)'),
              @nvarsccol = nullif(convert(nvarchar(MAX), @xmlsccol), '')

       SELECT @xmlcol = (SELECT lower(@tmp) AS Lågland
                         FOR XML RAW, ELEMENTS),
              @xmlsccol = (SELECT 1 AS Tag, NULL as Parent,
                                  upper(@tmpsc) AS [TÄST!1]
                           FOR    XML EXPLICIT)

       SELECT xmlcol, xmlsccol, nvarcol, nvarsccol
       FROM   xmltest
SQLEND

    if (not $output) {
       $spcode =~ s/\bOUTPUT\b//g;
    }

    sql($spcode);
}



#------------------------------------------------------------------------

sub datehash_compare {
  # Help routine to compare datehashes.
    my($val, $expect) = @_;

    foreach my $part (keys %$expect) {
       return 0 if not defined $$val{$part} or $$expect{$part} != $$val{$part};
    }
    return 1;
}


sub regional_to_ISO {
  # Help routine to convert date in regional to ISO
  my ($date) = @_;
  open DH, ">datehelperin.txt";
  print DH "$date\n";
  close DH;
  system("../helpers/datetesthelper");
  open DH, "datehelperout.txt";
  my $line = <DH>;
  close DH;
  my $ret = (split(/\s*£\s*/, $line))[1];
  $ret =~ s/^\s*|\s*$//g;
  return $ret;
}

sub open_testfile {
   open(TFILE, '>:utf8', TESTFILE);
   return \*TFILE;
}

sub get_testfile {
   open(TFILE, '<:utf8', TESTFILE);
   my $testfile = join('', <TFILE>);
   close TFILE;
   $testfile =~ s!\s*(\*/)?\ngo\s*$!\n!;
   return $testfile;
}

sub check_data {
   my ($checklogfile, $result, $params, $paramsbyref) = @_;

   my ($ix, $col, $valref, %filevalues);

   my $testfile;

   if ($checklogfile) {
      $testfile = get_testfile();
      if (not $params) {
         $testfile = get_testfile();
         $testfile =~ /\(([^\)]+)\)/;
         my $collist = $1;
         my @collist = split(/\s*,\s*/, $collist);
         unshift (@collist, undef);   # To make @collist 1-based.
         $testfile =~ s/\@P(\d+)\s*=/\@$collist[$1] =/g;
      }
   }

   foreach my $ix (0..$#tblcols) {
      my $col = $tblcols[$ix];
      next if not defined $col;

      my $valref;

      if (ref $params) {
         if (ref $params eq "ARRAY") {
            $valref = ($paramsbyref ? $$params[$ix] : \$$params[$ix]);
         }
         else {
            my $par = '@' . $col;
            $valref = ($paramsbyref ? $$params{$par} : \$$params{$par});
         }
      }
      else {
         $valref = undef;
      }

      my $resulttest = sprintf($test{$col}, '$$result{$col}', '$expectcol{$col}');
      my $paramtest  = sprintf($test{$col}, '$$valref', '$expectpar{$col}');
      my $comment    = defined $comment{$col} ? $comment{$col} : "";

      push(@testres,
           eval($resulttest) ? "ok %d" :
           "not ok %d # result '$col': <$$result{$col}>, expected: <$expectcol{$col}>" .
           "   $comment $@");
      if ($params and exists $expectpar{$col}) {
         push(@testres,
              eval($paramtest) ? "ok %d" :
              "not ok %d # param '$col': <$$valref>, expected: <$expectpar{$col}>  " .
              "    $comment $@");
      }

      if ($checklogfile) {
         my $filevalue;
         if ($testfile =~ m/\@$col\s*=\s*([^,\n]+)[,\n]/) {
            $filevalue = $1;
         }
         my $filetest   = ($filetest{$col} or '%s eq %s');
         $filetest = sprintf($filetest, '$filevalue', '$expectfile{$col}',
                                                      '$filevalue');
         push(@testres,
              eval($filetest) ? "ok %d" :
              "not ok %d # file '$col': <$filevalue>, expected: <$expectfile{$col}>" .
              "   $@");
     }
   }
}


sub do_tests {
    my ($X, $runlogfile, $typeclass, $testcase) = @_;

   $testcase = "<$typeclass" . (defined $testcase ? ", $testcase" : "") . ">";

   my ($result, @params, %params, @paramrefs, %paramrefs,
       @copy1, @copy2, $col);

   # Fill up parameter arrays. As the arrays are changed on each test,
   # fill up copies to refresh with as well.
   foreach $col (@tblcols) {
       if (defined $tbl{$col}) {
          push(@params, $tbl{$col});
          $params{'@' . $col} = $tbl{$col};
          push(@copy1, $tbl{$col});
          push(@copy2, $tbl{$col});
       }
       else {
          push(@params, undef);
          $params{'@' . $col} = undef;
          push(@copy1, undef);
          push(@copy2, undef);
       }
       push(@paramrefs,\$copy1[$#copy1]);
       $paramrefs{'@' . $col}    = \$copy2[$#copy2];
   }

   # Run test for combination.
   blurb("sql_sp $testcase unnamed params, no refs");
   $X->{LogHandle} = open_testfile();
   $result = sql_sp("${typeclass}_sp", \@params, HASH, SINGLEROW);
   undef $X->{LogHandle};
   check_data((not $runlogfile), $result, \@params, 0);

   if ($runlogfile) {
      blurb("Log file from sql_sp $testcase");
      my $logfile = get_testfile();
      $result = sql($logfile, HASH, SINGLEROW);
      check_data(0, $result, 0);
   }

   blurb("sql_sp $testcase named params, no refs");
   $result = sql_sp("${typeclass}_sp", \%params, HASH, SINGLEROW);
   undef $X->{LogHandle};
   check_data(0, $result, \%params, 0);

   blurb("sql_sp $testcase unnamed params, refs");
   $result = sql_sp("${typeclass}_sp", \@paramrefs, HASH, SINGLEROW);
   undef $X->{LogHandle};
   check_data(0, $result, \@paramrefs, 1);

   blurb("sql_sp $testcase named params, refs");
   $result = sql_sp("${typeclass}_sp", \%paramrefs, HASH, SINGLEROW);
   undef $X->{LogHandle};
   check_data(0, $result, \%paramrefs, 1);

   # Also test sql_insert.
   blurb("sql_insert $testcase");
   sql("TRUNCATE TABLE ${typeclass}");
   $X->{LogHandle} = open_testfile();
   sql_insert("${typeclass}", \%tbl);
   undef $X->{LogHandle};
   $result = sql("SELECT * FROM ${typeclass}", HASH, SINGLEROW);
   check_data((not $runlogfile), $result, 0);

   if ($runlogfile and $sqlver >= 7) {
      sql("TRUNCATE TABLE ${typeclass}");
      blurb("Log file from sql_insert $testcase");
      sql(get_testfile(), NORESULT);
      $result = sql("SELECT * FROM ${typeclass}", HASH, SINGLEROW);
      check_data(0, $result, 0);
   }

   $no_of_tests += 7 * scalar(keys %expectcol) +
                   4 * (scalar(keys %expectpar));

   if ($sqlver == 6 and $runlogfile) {
      $no_of_tests -= scalar(keys %expectcol);
   }
}



$^W = 1;
$| = 1;

$no_of_tests = 0;

my $X = testsqllogin();

$X->{'ErrInfo'}{RetStatOK}{4711}++;
$X->{'ErrInfo'}{NoWhine}++;
$X->{'ErrInfo'}{NeverPrint}{1708}++;  # Suppresses message for sql_variant table.

$sqlver = (split(/\./, $X->{SQL_version}))[0];

# Make sure that we have standard settings, except for ANSI_WARNINGS
# that we want to be off, as we test overlong input.
$X->sql(<<SQLEND);
SET ANSI_DEFAULTS ON
SET CURSOR_CLOSE_ON_COMMIT OFF
SET IMPLICIT_TRANSACTIONS OFF
SET ANSI_WARNINGS OFF
SQLEND

clear_test_data;
create_integer;

%tbl       = (intcol        =>   47114711,
              smallintcol   =>   -4711,
              tinyintcol    =>   111,
              floatcol      =>   123456789.456789,
              realcol       =>   123456789.456789,
              bitcol        =>   1);
%expectcol = (intcol        =>   $tbl{intcol} - 4711,
              smallintcol   =>   $tbl{smallintcol} - 4711,
              tinyintcol    =>   $tbl{tinyintcol} - 47,
              floatcol      =>   sprintf("%1.6f", $tbl{floatcol} - 4711),
              realcol       =>   $tbl{realcol} - 4711,
              bitcol        =>   ($tbl{bitcol} ? 0 : 1));
%expectpar = (intcol        =>   -2 * $tbl{intcol},
              smallintcol   =>   -2 * $tbl{smallintcol},
              tinyintcol    =>   2 * $tbl{tinyintcol},
              floatcol      =>    sprintf("%1.6f", -2 * $tbl{floatcol}),
              realcol       =>   -2 * $tbl{realcol},
              bitcol        =>   ($tbl{bitcol} ? 0 : 1));
%test      = (intcol        =>   '%s == %s',
              smallintcol   =>   '%s == %s',
              tinyintcol    =>   '%s == %s',
              floatcol      =>   'sprintf("%%1.6f", %s) eq %s',
              realcol       =>   'abs(%s - %s) < 10',
              bitcol        =>   '%s == %s');
do_tests($X, 1, 'integer', 'regular');

# Redo the tests, now will as many null values we can have.
%tbl       = (intcol        =>   undef,
              smallintcol   =>   undef,
              tinyintcol    =>   "087",
              floatcol      =>   undef,
              realcol       =>   undef,
              bitcol        =>   -1);
%expectcol = (intcol        =>   undef,,
              smallintcol   =>   undef,
              tinyintcol    =>   $tbl{tinyintcol} - 47,
              floatcol      =>   undef,
              realcol       =>   undef,
              bitcol        =>   ($tbl{bitcol} ? 0 : 1));
%expectpar = (intcol        =>   undef,
              smallintcol   =>   undef,
              tinyintcol    =>   2 * $tbl{tinyintcol},
              floatcol      =>   undef,
              realcol       =>   undef,
              bitcol        =>   ($tbl{bitcol} ? 0 : 1));
%test      = (intcol        =>   'not defined %s',
              smallintcol   =>   'not defined %s',
              tinyintcol    =>   '%s eq %s',
              floatcol      =>   'not defined %s',
              realcol       =>   'not defined %s',
              bitcol        =>   '%s == %s');
do_tests($X, 1, 'integer', 'null values');

drop_test_objects('integer');

#------------------------- CHARACTER --------------------------------
clear_test_data;
create_character;


%tbl       = (charcol      => "abc\x{00F6}",
              varcharcol   => "abc\x{010D}",
              varcharcol2  => "123456'8901234567890",
              textcol      => 'Hello Dolly! ' x 2000);
%expectcol = (charcol      => ' ' x 16 . "(\x{00F6}|o)" . 'cba',
              varcharcol   => "(\x{010D}|c)cba",
              varcharcol2  => "0987654321098'654321",
              textcol      => $tbl{textcol});
%expectpar = (charcol      => 'ABC' . "(\x{00D6}|O)" . ' ' x 16,
              varcharcol   => "ABC(\x{010C}|C)",
              varcharcol2  => $tbl{varcharcol2});
%test      = (charcol      => '%s =~ /^%s$/',
              varcharcol   => '%s =~ /^%s$/',
              varcharcol2  => ($sqlver > 6 ? '%s eq %s' : '%s =~ /^%s$/'),
              textcol      => '%s eq %s');
do_tests($X, 1, 'character');

# Known issue: NUL character in SQL command terminates command. This
# affects sql_insert on 6.5 and of this reason log file cannot be run.
%tbl       = (charcol      => '',
              varcharcol   => '',
              varcharcol2  => "123456789\x00123456789022",
              textcol      => '');
%expectcol = (charcol      => ' ' x 20,
              varcharcol   => ($sqlver == 6 ? ' ' : ''),
              varcharcol2  => ($sqlver > 6 ? "0987654321\x00987654321" :
                                             "(0987654321\x00)?987654321"),
              textcol      => ($sqlver == 6 ? ' ' : ''));
%expectpar = (charcol      => ' ' x 20,
              varcharcol   => ($sqlver == 6 ? ' ' : ''),
              varcharcol2  => substr($tbl{varcharcol2}, 0, 20));
%expectfile= (charcol      => ($sqlver == 6 ? "' ?'" : "''"),
              varcharcol   => ($sqlver == 6 ? "' ?'" : "''"),
              varcharcol2  => "'$tbl{varcharcol2}'",
              textcol     =>  ($sqlver == 6 ? "' ?'" : "''"));
%test      = (charcol      => '%s eq %s',
              varcharcol   => '%s eq %s',
              varcharcol2  => ($sqlver > 6 ? '%s eq %s' : '%s =~ /^%s$/'),
              textcol      => '%s eq %s');
%filetest  = (charcol      => ($sqlver > 6 ? '%s eq %s' : '%s =~ /^%s$/'),
              varcharcol   => ($sqlver > 6 ? '%s eq %s' : '%s =~ /^%s$/'),
              varcharcol2  => ($sqlver > 6 ? '%s eq %s' : '%s =~ /^%s$/'),
              textcol      => ($sqlver > 6 ? '%s eq %s' : '%s =~ /^%s$/'));
do_tests($X, 0, 'character', 'empty string');

# Known issue SQL7 (only) strips trailing blanks from varchar parameter.
%tbl       = (charcol      => undef,
              varcharcol   => undef,
              varcharcol2  => '  ',
              textcol      => undef);
%expectcol = (charcol      => undef,
              varcharcol   => undef,
              varcharcol2  => '  ',
              textcol      => undef);
%expectpar = (charcol      => undef,
              varcharcol   => undef,
              varcharcol2  => ($sqlver != 7 ? '  ' : '  ?'));
%test      = (charcol      => 'not defined %s',
              varcharcol   => 'not defined %s',
              varcharcol2  => ($sqlver != 7 ? '%s eq %s' : '%s =~ /^%s$/'),
              textcol      => 'not defined %s');
undef %filetest;
do_tests($X, 1, 'character', 'null');

drop_test_objects('character');

#------------------------- BINARY ---------------------------------
clear_test_data;
create_binary;

# Known issue: on 6.5 it appears that OUTPUT binary parameters loses
# trailing zero bytes.

#$X->{BinaryAsStr} = 1;    Default.
%tbl       = (bincol       => '4711ABCD',
              varbincol    => '4711ABCD',
              tstamp       => '0x00004711ABCD0009',
              imagecol     => '47119660AB002' x 10000);
%expectcol = (bincol       => '00' x 16 . 'CDAB1147',
              varbincol    => 'CDAB1147',
              tstamp       => '^[0-9A-F]{16}$',
              imagecol     => $tbl{'imagecol'});
%expectpar = (bincol       => '4711ABCD4711ABCD' .
                              ($sqlver > 6 ? '00' x 12 : ''),
              varbincol    => '4711ABCD4711ABCD',
              tstamp       => 'ABCD000900004711');
%test      = (bincol       => '%s eq %s',
              varbincol    => '%s eq %s',
              tstamp       => '%s =~ /%s/',
              imagecol     => '%s eq %s');
do_tests($X, 1, 'binary', 'BinaryAsStr = 1');

$X->{BinaryAsStr} = 1;
%tbl       = (bincol       => '0x',
              varbincol    => '0x',
              tstamp       => '0x',
              imagecol     => '0x');
%expectcol = (bincol       => '00' x 20,
              varbincol    => ($sqlver == 6 ? '00' : ''),
              tstamp       => '^[0-9A-F]{16}$',
              imagecol     => ($sqlver == 6 ? '00' : ''));
%expectpar = (bincol       => ($sqlver > 6 ? '00' x 20 : '0000'),
              varbincol    => ($sqlver == 6 ? '0000' : ''),
              tstamp       => '^' . '00' x 8 . '$');
%test      = (bincol       => '%s eq %s',
              varbincol    => '%s eq %s',
              tstamp       => '%s =~ /%s/',
              imagecol     => '%s eq %s');
do_tests($X, 1, 'binary', 'BinaryAsStr = 1 empty');


$X->{BinaryAsStr} = 'x';
# Known issue: SQL 7 appears to give wrong value back on 0x0000 for varbinpar.
%tbl       = (bincol       => '4711ABCD',
              varbincol    => '0x0000',
              tstamp       => '00004711ABCD0009',
              imagecol     => '47119660AB002' x 100);
%expectcol = (bincol       => '0x' . '00' x 16 . 'CDAB1147',
              varbincol    => '0x0000',
              tstamp       => '^0x[0-9A-F]{16}$',
              imagecol     => '0x' . $tbl{'imagecol'});
%expectpar = (bincol       => '0x4711ABCD4711ABCD' .
                              ($sqlver > 6 ? '00' x 12 : ''),
              varbincol    => ($sqlver != 7 ? '0x' . '00' x 4 : '0x00(000000)?'),
              tstamp       => '0xABCD000900004711');
%test      = (bincol       => '%s eq %s',
              varbincol    => ($sqlver != 7 ? '%s eq %s' : '%s =~ /^%s$/'),
              tstamp       => '%s =~ /%s/',
              imagecol     => '%s eq %s');
do_tests($X, 1, 'binary', 'BinaryAsStr = x');

$X->{BinaryAsStr} = 'x';
%tbl       = (bincol       => '',
              varbincol    => '',
              tstamp       => '0x',
              imagecol     => '');
%expectcol = (bincol       => '0x' . '00' x 20,
              varbincol    => '0x' . ($sqlver == 6 ? '00' : ''),
              tstamp       => '^0x[0-9A-F]{16}$',
              imagecol     => '0x' . ($sqlver == 6 ? '00' : ''));
%expectpar = (bincol       => '0x' . ($sqlver > 6 ? '00' x 20 : '0000'),
              varbincol    => '0x' . ($sqlver == 6 ? '0000' : ''),
              tstamp       => '^0x' . '00' x 8 . '$');
%test      = (bincol       => '%s eq %s',
              varbincol    => '%s eq %s',
              tstamp       => '%s =~ /%s/',
              imagecol     => '%s eq %s');
do_tests($X, 1, 'binary', 'BinaryAsStr = x. empty');


$X->{BinaryAsStr} = 0;
%tbl       = (bincol       => '4711ABCD',
              varbincol    => 'Typewriter',
              tstamp       => "\x00\x00/!#¤§=",
              imagecol     => 'Hello Dolly! ' x 10000);
%expectcol = (bincol       => "\x00" x 12 . 'DCBA1174',
              varbincol    => 'retirwepyT',
              tstamp       => "^(.|\\n){8}\$",
              imagecol     => $tbl{'imagecol'});
%expectpar = (bincol       => '47114711ABCD' .
                              ($sqlver > 6 ? "\x00" x 8 : ''),
              varbincol    => 'TypewriterTypewriter',
              tstamp       => "#¤§=\x00\x00/!");
%test      = (bincol       => '%s eq %s',
              varbincol    => '%s eq %s',
              tstamp       => '%s =~ /%s/',
              imagecol     => '%s eq %s');
do_tests($X, 1, 'binary', 'BinaryAsBinary');


%tbl       = (bincol       => '',
              varbincol    => '',
              tstamp       => '',
              imagecol     => '');
%expectcol = (bincol       => "\x00" x 20,
              varbincol    => ($sqlver == 6 ? "\x00" : ''),
              tstamp       => "^(.|\\n){8}\$",
              imagecol     => ($sqlver == 6 ? "\x00" : ''));
%expectpar = (bincol       => ($sqlver > 6 ? "\x00" x 20 : "\x00\x00"),
              varbincol    => ($sqlver == 6 ? "\x00\x00" : ''),
              tstamp       => '^' . "\x00" x 8 . '$');
%test      = (bincol       => '%s eq %s',
              varbincol    => '%s eq %s',
              tstamp       => '%s =~ /%s/',
              imagecol     => '%s eq %s');
do_tests($X, 1, 'binary', 'BinaryAsBinary, empty');

%tbl       = (bincol       => undef,
              varbincol    => undef,
              tstamp       => '00004711ABCD0009',
              imagecol     => undef);
%expectcol = (bincol       => undef,
              varbincol    => undef,
              tstamp       => "^(.|\\n){8}\$",
              imagecol     => undef);
%expectpar = (bincol       => undef,
              varbincol    => undef,
              tstamp       => '^47110000$');
%test      = (bincol       => 'not defined %s',
              varbincol    => 'not defined %s',
              tstamp       => '%s =~ /%s/',
              imagecol     => 'not defined %s');
do_tests($X, 1, 'binary', 'null');

drop_test_objects('binary');

#------------------------- DECIMAL --------------------------------
clear_test_data;
create_decimal;

#$X->{DecimalAsStr} = 0;   This should be default, so test this.
%tbl       = (deccol   => 123456912345678.456789,
              numcol   => 912345678.44,
              moneycol => 123456912345678.4567,
              dimecol  => 123456.4566);
%expectcol = (deccol   => $tbl{deccol}   - 12345678,
              numcol   => $tbl{numcol}   - 12345678,
              moneycol => $tbl{moneycol} - 12345678,
              dimecol  => $tbl{dimecol}  - 123456);
%expectpar = (deccol   => -2 * $tbl{deccol},
              numcol   => -$tbl{numcol} / 2,
              moneycol => -2 * $tbl{moneycol},
              dimecol  => -$tbl{dimecol} / 2);
%test      = (deccol   => 'abs(%s - %s) < 100',
              numcol   => 'abs(%s - %s) < 1E-6',
              moneycol => 'abs(%s - %s) < 100',
              dimecol  => 'abs(%s - %s) < 1E-6');
do_tests($X, 1, 'decimal', 'DecimalAsStr = 0');


$X->{DecimalAsStr} = 1; # Input is still numeric.
%tbl       = (deccol   => 123456912345678.456789,
              numcol   => 912345678.44,
              moneycol => 123456912345678.4567,
              dimecol  => 123456.4566);
%expectcol = (deccol   => $tbl{deccol}   - 12345678,
              numcol   => '900000000.44',
              moneycol => $tbl{moneycol} - 12345678,
              dimecol  => '0.4566');
%expectpar = (deccol   => -2 * $tbl{deccol},
              numcol   => '-456172839.22',
              moneycol => -2 * $tbl{moneycol},
              dimecol  => '-61728.2283');
%test      = (deccol   => 'abs(%s - %s) < 100',
              numcol   => '%s eq %s',
              moneycol => 'abs(%s - %s) < 100',
              dimecol  => '%s eq %s');
do_tests($X, 1, 'decimal', 'DecimalAsStr = 1, num in');


# Now we also send strings in.
%tbl       = (deccol   => '123456912345678.456789',
              numcol   => '912345678.44',
              moneycol => '123456912345678.4567',
              dimecol  => '123456.4566');
%expectcol = (deccol   => '123456900000000.456789',
              numcol   => '900000000.44',
              moneycol => '123456900000000.4567',
              dimecol  => '0.4566');
%expectpar = (deccol   => '-246913824691356.913578',
              numcol   => '-456172839.22',
              moneycol => '-246913824691356.9134',
              dimecol  => '-61728.2283');
%test      = (deccol   => '%s eq %s',
              numcol   => '%s eq %s',
              moneycol => '%s eq %s',
              dimecol  => '%s eq %s');
do_tests($X, 1, 'decimal', 'DecimalAsStr = 1, str in');

# And test null values.
%tbl       = (deccol   => undef,
              numcol   => undef,
              moneycol => undef,
              dimecol  => undef);
%expectcol = (deccol   => undef,
              numcol   => undef,
              moneycol => undef,
              dimecol  => undef);
%expectpar = (deccol   => undef,
              numcol   => undef,
              moneycol => undef,
              dimecol  => undef);
%test      = (deccol   => 'not defined %s',
              numcol   => 'not defined %s',
              moneycol => 'not defined %s',
              dimecol  => 'not defined %s');
do_tests($X, 1, 'decimal', 'null values');

drop_test_objects('decimal');

#------------------------- DATETIME --------------------------------
clear_test_data;
create_datetime;

# For datetime we must read the log file for most cases, since most date
# strings will only be exuectable with some dateformat settings - or
# even not at all.

#$X->{DateimeOption} = DATETIME_ISO    -- The default.
%tbl       = (datetimecol  => '1996-08-13 04:36:24.997',
              smalldatecol => '1996-08-13 04:36');
%expectcol = (datetimecol  => '1996-08-30 04:36:24.997',
              smalldatecol => '1996-11-13 04:36');
%expectpar = (datetimecol  => '1996-08-13 08:36:24.997',
              smalldatecol => '1996-08-13 04:50');
%expectfile= (datetimecol  => "'1996-08-13 04:36:24.997'",
              smalldatecol => "'1996-08-13 04:36'");
%test      = (datetimecol  => '%s eq %s',
              smalldatecol => '%s eq %s');
do_tests($X, 0, 'datetime', 'ISO in/out');

%tbl       = (datetimecol  => undef,
              smalldatecol => undef);
%expectcol = (datetimecol  => undef,
              smalldatecol => undef);
%expectpar = (datetimecol  => undef,
              smalldatecol => undef);
undef %expectfile;
%test      = (datetimecol  => 'not defined %s',
              smalldatecol => 'not defined %s');
do_tests($X, 1, 'datetime', 'ISO in/out, nulls');


%tbl       = (datetimecol  => '1996-08-13',
              smalldatecol => '1996-08-13');
%expectcol = (datetimecol  => '1996-08-30 00:00:00.000',
              smalldatecol => '1996-11-13 00:00');
%expectpar = (datetimecol  => '1996-08-13 04:00:00.000',
              smalldatecol => '1996-08-13 00:14');
%expectfile= (datetimecol  => "'1996-08-13'",
              smalldatecol => "'1996-08-13'");
%test      = (datetimecol  => '%s eq %s',
              smalldatecol => '%s eq %s');
do_tests($X, 0, 'datetime', 'ISO dates only');

%tbl       = (datetimecol  => '19960813 04:36:24.997',
              smalldatecol => '19960813 04:36');
%expectcol = (datetimecol  => '1996-08-30 04:36:24.997',
              smalldatecol => '1996-11-13 04:36');
%expectpar = (datetimecol  => '1996-08-13 08:36:24.997',
              smalldatecol => '1996-08-13 04:50');
undef %expectfile;   # The log file can be used, hooray!
%test      = (datetimecol  => '%s eq %s',
              smalldatecol => '%s eq %s');
do_tests($X, 1, 'datetime', 'YYYYMMDD in/ISO out');

%tbl       = (datetimecol  => '19960813',
              smalldatecol => '19960813');
%expectcol = (datetimecol  => '1996-08-30 00:00:00.000',
              smalldatecol => '1996-11-13 00:00');
%expectpar = (datetimecol  => '1996-08-13 04:00:00.000',
              smalldatecol => '1996-08-13 00:14');
undef %expectfile;
%test      = (datetimecol  => '%s eq %s',
              smalldatecol => '%s eq %s');
do_tests($X, 1, 'datetime', 'YYYMMDD only in/ISO out');

%tbl       = (datetimecol  => '1994-08-13Z',
              smalldatecol => '1994-08-13Z');
%expectcol = (datetimecol  => '1994-08-30 00:00:00.000',
              smalldatecol => '1994-11-13 00:00');
%expectpar = (datetimecol  => '1994-08-13 04:00:00.000',
              smalldatecol => '1994-08-13 00:14');
if ($sqlver >= 9) {
   undef %expectfile;
}
else {
   %expectfile= (datetimecol  => "'1994-08-13Z'",
                 smalldatecol => "'1994-08-13Z'");
}
%test      = (datetimecol  => '%s eq %s',
              smalldatecol => '%s eq %s');
do_tests($X, ($sqlver >= 9), 'datetime', 'YYYY-MM-DDZ');


%tbl       = (datetimecol  => '1996-08-13T04:36:24.997',
              smalldatecol => '1996-08-13T04:36');
%expectcol = (datetimecol  => '1996-08-30 04:36:24.997',
              smalldatecol => '1996-11-13 04:36');
%expectpar = (datetimecol  => '1996-08-13 08:36:24.997',
              smalldatecol => '1996-08-13 04:50');
%expectfile= (datetimecol  => "'1996-08-13T04:36:24.997'",
              smalldatecol => "'1996-08-13T04:36'");
%test      = (datetimecol  => '%s eq %s',
              smalldatecol => '%s eq %s');
do_tests($X, 0, 'datetime', 'XML in/ ISO out');

%tbl       = (datetimecol  => {Year => 1996, Month => 8, Day => 13,
                               Hour => 4, Minute => 36, Second => 24,
                               Fraction => 997},
              smalldatecol => {Year => 1996, Month => 8, Day => 13,
                               Hour => 4, Minute => 36});
%expectcol = (datetimecol  => '1996-08-30 04:36:24.997',
              smalldatecol => '1996-11-13 04:36');
%expectpar = (datetimecol  => '1996-08-13 08:36:24.997',
              smalldatecol => '1996-08-13 04:50');
%test      = (datetimecol  => '%s eq %s',
              smalldatecol => '%s eq %s');
%expectfile= (datetimecol  => "^'HASH\\(",
              smalldatecol => "^'HASH\\(");
%filetest  = (datetimecol  => '%s =~ /%s/',
              smalldatecol => '%s =~ /%s/');
do_tests($X, 0, 'datetime', 'Hash in, ISO out');

%tbl       = (datetimecol  => {Year => 1996, Month => 8, Day => 13},
              smalldatecol => {Year => 1996, Month => 8, Day => 13});
%expectcol = (datetimecol  => '1996-08-30 00:00:00.000',
              smalldatecol => '1996-11-13 00:00');
%expectpar = (datetimecol  => '1996-08-13 04:00:00.000',
              smalldatecol => '1996-08-13 00:14');
%test      = (datetimecol  => '%s eq %s',
              smalldatecol => '%s eq %s');
%expectfile= (datetimecol  => "^'HASH\\(",
              smalldatecol => "^'HASH\\(");
%filetest  = (datetimecol  => '%s =~ /%s/',
              smalldatecol => '%s =~ /%s/');
do_tests($X, 0, 'datetime', 'Hash in dates only');

%tbl       = (datetimecol  => 3.25,
              smalldatecol => 4);
%expectcol = (datetimecol  => '1900-01-19 06:00:00.000',
              smalldatecol => '1900-04-03 00:00');
%expectpar = (datetimecol  => '1900-01-02 10:00:00.000',
              smalldatecol => '1900-01-03 00:14');
%expectfile= %tbl;
%test      = (datetimecol  => '%s eq %s',
              smalldatecol => '%s eq %s');
do_tests($X, 0, 'datetime', 'Float in/ ISO out');

# For regional settings we need a help file.
{
   open DH, ">datehelperin.txt";
   print DH "1996-08-13 04:36:24", "\n", "1996-08-13 04:36", "\n";
   close DH;
   system("../helpers/datetesthelper");
   open DH, "datehelperout.txt";
   my $line1 = <DH>;
   my $line2 = <DH>;
   close DH;
   my ($regionaldate) = split(/\s*£\s*/, $line1);
   my ($regionalsmall) = split(/\s*£\s*/, $line2);

   %tbl       = (datetimecol  => $regionaldate,
                 smalldatecol => $regionalsmall);
   %expectcol = (datetimecol  => '1996-08-30 04:36:24.000',
                 smalldatecol => '1996-11-13 04:36');
   %expectpar = (datetimecol  => '1996-08-13 08:36:24.000',
                 smalldatecol => '1996-08-13 04:50');
   %expectfile= (datetimecol  => "'$regionaldate'",
                 smalldatecol => "'$regionalsmall'");
   %test      = (datetimecol  => '%s eq %s',
                 smalldatecol => '%s eq %s');
   do_tests($X, 0, 'datetime', 'Reg setting long in/ISO out');
}

$X->{DatetimeOption} = DATETIME_STRFMT;
%tbl       = (datetimecol  => '19960813 04:36:24.997',
              smalldatecol => '19960813 04:36');
%expectcol = (datetimecol  => '19960830 04:36:24.997',
              smalldatecol => '19961113 04:36(:00)?');
%expectpar = (datetimecol  => '19960813 08:36:24.997',
              smalldatecol => '19960813 04:50(:00)?');
undef %expectfile;
%test      = (datetimecol  => '%s eq %s',
              smalldatecol => '%s =~ /^%s$/');
do_tests($X, 1, 'datetime', 'ISO in/ STRFMT out default');

$X->{DateFormat} = "%d.%m.%y";
undef $X->{msecFormat};
%tbl       = (datetimecol  => '19960813 04:36:24.997',
              smalldatecol => '19960813 04:36');
%expectcol = (datetimecol  => '30.08.96',
              smalldatecol => '13.11.96');
%expectpar = (datetimecol  => '13.08.96',
              smalldatecol => '13.08.96');
undef %expectfile;
%test      = (datetimecol  => '%s eq %s',
              smalldatecol => '%s eq %s');
do_tests($X, 1, 'datetime', 'ISO in/ STRFMT out custom');

$X->{DatetimeOption} = DATETIME_FLOAT;
%tbl       = (datetimecol  => '19000102 06:00',
              smalldatecol => '19000104');
%expectcol = (datetimecol  => 20.25,
              smalldatecol => 95);
%expectpar = (datetimecol  => 3 + 10/24,
              smalldatecol => 5 + 14/(24*60));
undef %expectfile;
%test      = (datetimecol  => 'abs(%s - %s) < 1E-9',
              smalldatecol => 'abs(%s - %s) < 1E-9');
do_tests($X, 1, 'datetime', 'ISO in/ FLOAT out');

$X->{DatetimeOption} = DATETIME_HASH;
%tbl       = (datetimecol  => '19960813 04:36:24.997',
              smalldatecol => '19960813 04:36');
%expectcol = (datetimecol  => {Year => 1996, Month => 8, Day => 30,
                               Hour => 4, Minute => 36, Second => 24,
                               Fraction => 997},
              smalldatecol => {Year => 1996, Month => 11, Day => 13,
                               Hour => 4, Minute => 36});
%expectpar = (datetimecol  => {Year => 1996, Month => 8, Day => 13,
                               Hour => 8, Minute => 36, Second => 24,
                               Fraction => 997},
              smalldatecol => {Year => 1996, Month => 8, Day => 13,
                               Hour => 4, Minute => 50});
undef %expectfile;
%test      = (datetimecol  => 'datehash_compare(%s, %s)',
              smalldatecol => 'datehash_compare(%s, %s)');
do_tests($X, 1, 'datetime', 'ISO in/HASH out');

%tbl       = (datetimecol  => undef,
              smalldatecol => undef);
%expectcol = (datetimecol  => undef,
              smalldatecol => undef);
%expectpar = (datetimecol  => undef,
              smalldatecol => undef);
undef %expectfile;
%test      = (datetimecol  => 'not defined %s',
              smalldatecol => 'not defined %s');
do_tests($X, 1, 'datetime', 'NULL in/hash out');

$X->{DatetimeOption} = DATETIME_REGIONAL;
%tbl       = (datetimecol  => '19960813 04:36:24',
              smalldatecol => '19960813 04:36');
%expectcol = (datetimecol  => '1996-08-30 04:36:24',
              smalldatecol => '1996-11-13 04:36(:00)?');
%expectpar = (datetimecol  => '1996-08-13 08:36:24',
              smalldatecol => '1996-08-13 04:50(:00)?');
undef %expectfile;
%test      = (datetimecol  => 'regional_to_ISO(%s) eq %s',
              smalldatecol => 'regional_to_ISO(%s) =~ /^%s$/');
do_tests($X, 1, 'datetime', 'ISO in/ REGIONAL out');

drop_test_objects('datetime');

#------------------------- GUID + NULLBIT-------------------------------
# From here we're SQL7 and up only.
goto finally if $sqlver == 6;

clear_test_data;
create_guid;

%tbl       = (guidcol     => 'FF0DCAF3-CFFC-4C9B-AE4B-C08B2000871C',
              nullbitcol  => 1);
%expectcol = (guidcol     => '{000DCA03-C00C-4C9B-AE4B-C08B2000871C}',
              nullbitcol  => 0);
%expectpar = (guidcol     => '{AA0DCAA3-CAAC-4C9B-AE4B-C08B2000871C}',
              nullbitcol  => 0);
%test      = (guidcol     => '%s eq %s',
              nullbitcol  => '%s eq %s');
do_tests($X, 1, 'guid', 'unbraced');

%tbl       = (guidcol     => '{FF0DCAF3-CFFC-4C9B-AE4B-C08B2000871C}',
              nullbitcol  => 0);
%expectcol = (guidcol     => '{000DCA03-C00C-4C9B-AE4B-C08B2000871C}',
              nullbitcol  => 1);
%expectpar = (guidcol     => '{AA0DCAA3-CAAC-4C9B-AE4B-C08B2000871C}',
              nullbitcol  => 1);
%test      = (guidcol     => '%s eq %s',
              nullbitcol  => '%s eq %s');
do_tests($X, 1, 'guid', 'braced');

%tbl       = (guidcol     => undef,
              nullbitcol  => undef);
%expectcol = (guidcol     => undef,
              nullbitcol  => undef);
%expectpar = (guidcol     => undef,
              nullbitcol  => undef);
%test      = (guidcol     => 'not defined %s',
              nullbitcol  => 'not defined %s');
do_tests($X, 1, 'guid', 'null values');

drop_test_objects('guid');

#------------------------- UNICODE --------------------------------
clear_test_data;
create_unicode;

my $nvarcharcol = "\x{0144}varcharcol";
my $ncharcol2 = 'nchärcöl2';
binmode(STDOUT, ':utf8:');

%tbl       = (ncharcol      => "\x{00E6}\x{00E5}\x{00F6}\x{FFFD}",
              $nvarcharcol  => "abc\x{0157}",
              $ncharcol2    => "123456'890123456789\x{010B}",
              ntextcol      => '21 pa\x{017A}dziernika 2004 ' x 2000);
%expectcol = (ncharcol      => ' ' x 16 . "\x{FFFD}\x{00F6}\x{00E5}\x{00E6}",
              $nvarcharcol  => "\x{0157}cba",
              $ncharcol2    => "\x{010B}987654321098'654321",
              ntextcol      => $tbl{ntextcol});
%expectpar = (ncharcol      => "\x{00C6}\x{00C5}\x{00D6}\x{FFFD}" . ' ' x 16,
              $nvarcharcol  => "ABC\x{0156}",
              $ncharcol2    => "123456'890123456789\x{010A}");
%test      = (ncharcol      => '%s eq %s',
              $nvarcharcol  => '%s eq %s',
              $ncharcol2    => '%s eq %s',
              ntextcol      => '%s eq %s');
do_tests($X, 1, 'unicode');

# Known issue: NULL terminates strings in literal SQL commands, so log
# file cannot be used. Unknown if this is an SQL Server bug.
%tbl       = (ncharcol      => '',
              $nvarcharcol  => '',
              $ncharcol2    => "\x001234567890'23456789022",
              ntextcol      => '');
%expectcol = (ncharcol      => ' ' x 20,
              $nvarcharcol  => '',
              $ncharcol2    => "98765432'0987654321\x00",
              ntextcol      => '');
%expectpar = (ncharcol      => ' ' x 20,
              $nvarcharcol  => '',
              $ncharcol2    => "\x001234567890'23456789");
%expectfile= (ncharcol      => "N''",
              $nvarcharcol  => "N''",
              $ncharcol2    => "N'\x001234567890''23456789022'",
              ntextcol      => "N''");
%test      = (ncharcol      => '%s eq %s',
              $nvarcharcol  => '%s eq %s',
              $ncharcol2    => '%s eq %s',
              ntextcol      => '%s eq %s');
do_tests($X, 0, 'unicode', 'empty string');

# Known issue SQL7 (only) strips trailing blanks from nvarchar parameter.
%tbl       = (ncharcol      => undef,
              $nvarcharcol  => undef,
              $ncharcol2    => '  ',
              ntextcol      => undef);
%expectcol = (ncharcol      => undef,
              $nvarcharcol  => undef,
              $ncharcol2    => ' ' x 20,
              ntextcol      => undef);
%expectpar = (ncharcol      => undef,
              $nvarcharcol  => undef,
              $ncharcol2    => ' ' x 20);
%test      = (ncharcol      => 'not defined %s',
              $nvarcharcol  => 'not defined %s',
              $ncharcol2    => '%s eq %s',
              ntextcol      => 'not defined %s');
do_tests($X, 1, 'unicode', 'null');

drop_test_objects('unicode');

#------------------------- BIGINT -------------------------------
# From here we're SQL 2000 and up only.
goto finally if $sqlver == 7;

#------------------------- BIGINT --------------------------------
clear_test_data;
create_bigint;

$X->{DecimalAsStr} = 0;
%tbl       = (bigintcol   => 123456912345678);
%expectcol = (bigintcol   => $tbl{bigintcol} - 12345678);
%expectpar = (bigintcol   => -2 * $tbl{bigintcol});
%test      = (bigintcol   => 'abs(%s - %s) < 100');
do_tests($X, 1, 'bigint', 'DecimalAsStr = 0');

$X->{DecimalAsStr} = 1; # Input is still numeric.
%tbl       = (bigintcol   => 123456912345678);
%expectcol = (bigintcol   => $tbl{bigintcol} - 12345678);
%expectpar = (bigintcol   => -2 * $tbl{bigintcol});
%test      = (bigintcol   => 'abs(%s - %s) < 100');
do_tests($X, 1, 'bigint', 'DecimalAsStr = 1, num in');

# Now we also send strings in.
%tbl       = (bigintcol   => '123456912345678');
%expectcol = (bigintcol   => '123456900000000');
%expectpar = (bigintcol   => '-246913824691356');
%test      = (bigintcol   => '%s eq %s');
do_tests($X, 1, 'bigint', 'DecimalAsStr = 1, str in');

# And test null values.
%tbl       = (bigintcol => undef);
%expectcol = (bigintcol => undef);
%expectpar = (bigintcol => undef);
%test      = (bigintcol => 'not defined %s');
do_tests($X, 1, 'bigint', 'null values');

drop_test_objects('bigint');

#---------------------------- SQL_VARIANT ------------------------------
clear_test_data;
create_sql_variant;

# Test send in outtype to tell how data is to be returned. intype is the
# base type for the expression for the inparameter.

# This is always the same, because we never run the log file. Parameter
# should always be an nvarchar constant.
%filetest  = (varcol  => '%s eq %s');

# Note here that the test for outtype is really a dummy type - this is not
# an output parameter. But this is how the framework works.
%tbl       = (varcol  => 112,
              intype  => undef,
              outtype => 'bit');
%expectcol = (varcol  => 1,
              intype  => 'int',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => 1,
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s == %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'bit');

%tbl       = (varcol  => 112,
              intype  => undef,
              outtype => 'tinyint');
%expectcol = (varcol  => 62,
              intype  => 'int',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => 224,
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s == %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'tinyint');

%tbl       = (varcol  => -10112,
              intype  => undef,
              outtype => 'smallint');
%expectcol = (varcol  => -10162,
              intype  => 'int',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => 20224,
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s == %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'int');

%tbl       = (varcol  => 1120000,
              intype  => undef,
              outtype => 'int');
%expectcol = (varcol  => 1119950,
              intype  => 'int',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => -2240000,
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%test      = (varcol  => '%s == %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
do_tests($X, 0, 'sql_variant', 'int');

$X->{DecimalAsStr} = 0;
%tbl       = (varcol  => 123456912345678,
              intype  => undef,
              outtype => 'bigint');
%expectcol = (varcol  => 123456900000000,
              intype  => 'float',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => -246913824691356,
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => 'abs(%s - %s) < 100',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'bigint');

$X->{DecimalAsStr} = 1;
%tbl       = (varcol  => '123456912345678',
              intype  => undef,
              outtype => 'bigint');
%expectcol = (varcol  => '123456900000000',
              intype  => 'varchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => '-246913824691356',
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s eq %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'bigint as str');

%tbl       = (varcol  => 786.987,
              intype  => undef,
              outtype => 'real');
%expectcol = (varcol  => 736.987,
              intype  => 'float',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => -1573.974,
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => 'abs(%s - %s) < 0.01',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'real');

%tbl       = (varcol  => -786.987,
              intype  => undef,
              outtype => 'float');
%expectcol = (varcol  => -836.987,
              intype  => 'float',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => 1573.974,
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => 'abs(%s - %s) < 1E-7',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'float');

$X->{DecimalAsStr} = 0;
%tbl       = (varcol  => -912345678.12,
              intype  => undef,
              outtype => 'numeric');
%expectcol = (varcol  => -924691356.12,
              intype  => 'float',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => 1824691356.24,
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => 'abs(%s - %s) < 0.001',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'numeric, dec as num');

$X->{DecimalAsStr} = 1;
%tbl       = (varcol  => '123456912345678.123456',
              intype  => undef,
              outtype => 'decimal');
%expectcol = (varcol  => '123456900000000.123456',
              intype  => 'varchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => '-246913824691356.246912',
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s eq %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'decimal as str');

$X->{DecimalAsStr} = 1;
%tbl       = (varcol  => '12345.3412',
              intype  => undef,
              outtype => 'smallmoney');
%expectcol = (varcol  => '0.3412',
              intype  => 'varchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => '-24690.6824',
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s eq %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'smallmoney as str');

$X->{DecimalAsStr} = 0;
%tbl       = (varcol  => '123456912345678.123456',
              intype  => undef,
              outtype => 'decimal');
%expectcol = (varcol  => 123456900000000.123456,
              intype  => 'varchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => -246913824691356.246912,
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => 'abs(%s - %s) < 0.01',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'money as dec');

$X->{DatetimeOption} = DATETIME_ISO;
%tbl       = (varcol  => {Year => 1996, Month => 10, Day => 21,
                          Hour => 14, Minute => 16, Second => 23},
              intype  => undef,
              outtype => 'datetime');
%expectcol = (varcol  => '1996-09-01 14:16:23.000',
              intype  => 'datetime',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => '1996-10-22 00:16:23.000',
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s eq %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'datetime hash/iso');

%tbl       = (varcol  => {Year => 1996, Month => 10, Day => 21,
                          Hour => 14, Minute => 16, Second => 23},
              intype  => undef,
              outtype => 'smalldatetime');
%expectcol = (varcol  => '1996-09-01 14:16',
              intype  => 'datetime',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => '1996-10-22 00:16',
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s eq %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'smalldatetime hash/iso');

$X->{DatetimeOption} = DATETIME_REGIONAL;
%tbl       = (varcol  => {Year => 1996, Month => 10, Day => 21,
                          Hour => 14, Minute => 16, Second => 23},
              intype  => undef,
              outtype => 'datetime');
%expectcol = (varcol  => '1996-09-01 14:16:23',
              intype  => 'datetime',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => '1996-10-22 00:16:23',
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => 'regional_to_ISO(%s) eq %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'datetime hash/regional');

$X->{DatetimeOption} = DATETIME_HASH;
%tbl       = (varcol  => '19961021 14:16:23',
              intype  => undef,
              outtype => 'datetime');
%expectcol = (varcol  => {Year => 1996, Month => 9, Day => 1, Hour => 14,
                          Minute => 16, Second => 23, Fraction => 0},
              intype  => 'varchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => {Year => 1996, Month => 10, Day => 22, Hour => 0,
                          Minute => 16, Second => 23, Fraction => 0},
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => 'datehash_compare(%s, %s)',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'datetime iso/hash');

%tbl       = (varcol  => '19961021 14:16',
              intype  => undef,
              outtype => 'smalldatetime');
%expectcol = (varcol  => {Year => 1996, Month => 9, Day => 1, Hour => 14,
                          Minute => 16},
              intype  => 'varchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => {Year => 1996, Month => 10, Day => 22, Hour => 0,
                          Minute => 16},
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => 'datehash_compare(%s, %s)',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'smalldatetime iso/hash');

%tbl       = (varcol  => "abc",
              intype  => undef,
              outtype => 'char');
%expectcol = (varcol  => ' ' x 17 . "cba",
              intype  => 'varchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => "ABC" . ' ' x 17,
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s eq %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'char');

%tbl       = (varcol  => "123456789\x00123456789nn",
              intype  => undef,
              outtype => 'varchar');
%expectcol = (varcol  => "n987654321\x00987654321",
              intype  => 'varchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => "123456789\x00123456789N",
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s eq %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'varchar');

%tbl       = (varcol  => "abc\x{010B}\x{FFFD}",
              intype  => undef,
              outtype => 'nchar');
%expectcol = (varcol  => ' ' x 15 . "\x{FFFD}\x{010B}cba",
              intype  => 'nvarchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => "ABC\x{010A}\x{FFFD}" . ' ' x 15,
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s eq %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'nchar');

%tbl       = (varcol  => "\x{010B}123456789\x{FFFD}",
              intype  => undef,
              outtype => 'nvarchar');
%expectcol = (varcol  => "\x{FFFD}987654321\x{010B}",
              intype  => 'nvarchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => "\x{010A}123456789\x{FFFD}",
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s eq %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'nvarchar');

$X->{BinaryAsStr} = 0;
%tbl       = (varcol  => "123456789\x{FFFD}",
              intype  => undef,
              outtype => 'binary');
%expectcol = (varcol  => "1\x002\x003\x004\x005\x006\x007\x008\x009\x00\xFD\xFF",
              intype  => 'nvarchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => "1\x002\x003\x004\x005\x006\x007\x008\x009\x00\xFD\xFF",
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s eq %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'binary as bin');

$X->{BinaryAsStr} = 1;
%tbl       = (varcol  => "abc",
              intype  => undef,
              outtype => 'varbinary');
%expectcol = (varcol  => "616263",
              intype  => 'varchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => "616263",
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s eq %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'varbinary as str');

$X->{BinaryAsStr} = 'x';
%tbl       = (varcol  => "abc",
              intype  => undef,
              outtype => 'binary');
%expectcol = (varcol  => "0x616263" . '00' x 17,
              intype  => 'varchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => "0x616263" . '00' x 17,
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s eq %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'binary as 0x');

%tbl       = (varcol  => "1B2EA68F-6E22-4471-B67E-2E4EFCC283CD",
              intype  => undef,
              outtype => 'uniqueidentifier');
%expectcol = (varcol  => "{1B2EA68F-6E22-4471-B67E-2E4EFCC283CD}",
              intype  => 'varchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => "{1B2EA68F-6E22-4471-B67E-2E4EFCC283CD}",
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => '%s eq %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'uniqueidentifier');

%tbl       = (varcol  => [9878],
              intype  => undef,
              outtype => 'NULL');
%expectcol = (varcol  => undef,
              intype  => 'varchar',
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => undef,
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "N'$tbl{'varcol'}'",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => 'not defined %s',
              intype  => '%s eq %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'NULL out');

%tbl       = (varcol  => undef,
              intype  => undef,
              outtype => 'datetime');
%expectcol = (varcol  => undef,
              intype  => undef,
              outtype => $tbl{'outtype'});
%expectpar = (varcol  => undef,
              intype  => $expectcol{'intype'},
              outtype => $tbl{'outtype'});
%expectfile= (varcol  => "NULL",
              intype  => 'NULL',
              outtype => "N'$tbl{'outtype'}'");
%test      = (varcol  => 'not defined %s',
              intype  => 'not defined %s',
              outtype => '%s eq %s');
do_tests($X, 0, 'sql_variant', 'NULL in/out');

drop_test_objects('sql_variant');

#-------------------------- (N)VARCHAR MAX -----------------------------
# From here we're SQL 2005 and up only.
goto finally if $sqlver == 8;

clear_test_data;
create_varcharmax;

# When we run with SQLOLEDB, the (MAX) will be passed forth and back
# as (8000) or (4000).
%tbl       = (varcharcol   => 'Hello Dolly! ' x 2000,
              nvarcharcol  => "21 pa\x{017A}dziernika 2004 " x 2000);
if ($X->{Provider} == PROVIDER_SQLNCLI) {
   %expectcol = (varcharcol  => ' !ylloD olleH' x 2000,
                 nvarcharcol => " 4002 akinreizd\x{017A}ap 12" x 2000);
   %expectpar = (varcharcol  => 'HELLO DOLLY! ' x 2000,
                 nvarcharcol => "21 PA\x{0179}DZIERNIKA 2004 " x 2000);
   %test      = (varcharcol  => '%s eq %s',
                 nvarcharcol => '%s eq %s');
}
else {
   %expectcol = (varcharcol   => '('   . 'olleH' . ' !ylloD olleH' x 615 .
                                 ')|(' . ' !ylloD olleH' x 2000 . ')',
                 nvarcharcol  => '(' . "eizd\x{017A}ap 12" .
                                       " 4002 akinreizd\x{017A}ap 12" x 190 .
                                 ')|(' . " 4002 akinreizd\x{017A}ap 12" x 2000 . ')');
   %expectpar = (varcharcol   => 'HELLO DOLLY! ' x 615 . 'HELLO',
                 nvarcharcol  => "21 PA\x{0179}DZIERNIKA 2004 " x 190 .
                                 "21 PA\x{0179}DZIE");
   %test      = (varcharcol   => '%s =~ /^%s$/',
                 nvarcharcol  => '%s =~ /^%s$/');
}
do_tests($X, 1, 'varcharmax');


%tbl       = (varcharcol   => '',
              nvarcharcol  => '');
%expectcol = (varcharcol   => '',
              nvarcharcol  => '');
%expectpar = (varcharcol   => '',
              nvarcharcol  => '');
%test      = (varcharcol   => '%s eq %s',
              nvarcharcol  => '%s eq %s');
do_tests($X, 1, 'varcharmax', 'empty string');

%tbl       = (varcharcol   => undef,
              nvarcharcol  => '   ');
%expectcol = (varcharcol   => undef,
              nvarcharcol  => '   ');
%expectpar = (varcharcol   => undef,
              nvarcharcol  => '   ');
%test      = (varcharcol   => 'not defined %s',
              nvarcharcol  => '%s eq %s');
do_tests($X, 1, 'varcharmax', 'null');

drop_test_objects('varcharmax');

#-------------------------- (N)VARBINARY MAX -----------------------------

clear_test_data;
create_varbinmax;

# When we run with SQLOLEDB, the (MAX) will be passed forth and back
# as (8000) or (4000).
$X->{BinaryAsStr} = 1;
%tbl       = (varbincol    => '47119660AB0102' x 10000);
if ($X->{Provider} == PROVIDER_SQLNCLI) {
   %expectcol = (varbincol    => '0201AB60961147' x 10000);
   %expectpar = (varbincol    => '47119660AB0102' x 20000);
   %test      = (varbincol    => '%s eq %s');
}
else {
   %expectcol = (varbincol    => '(' . '01AB60961147' . '0201AB60961147' x 1142 .
                                 ')|(' . '0201AB60961147' x 10000 . ')');
   %expectpar = (varbincol    => '47119660AB0102' x 1142 . '47119660AB01');
   %test      = (varbincol    => '%s =~ /^%s$/');
}
do_tests($X, 1, 'varbinmax', 'as str');

%tbl       = (varbincol    => '0x');
%expectcol = (varbincol    => '');
%expectpar = (varbincol    => '');
%test      = (varbincol    => '%s eq %s');
do_tests($X, 1, 'varbinmax', 'empty string');

$X->{BinaryAsStr} = 'x';
%tbl       = (varbincol    => '0x' . '47119660AB0102' x 10000);
if ($X->{Provider} == PROVIDER_SQLNCLI) {
   %expectcol = (varbincol    => '0x' . '0201AB60961147' x 10000);
   %expectpar = (varbincol    => '0x' . '47119660AB0102' x 20000);
   %test      = (varbincol    => '%s eq %s');
}
else {
   %expectcol = (varbincol    => '0x' .
                                 '(' . '01AB60961147' . '0201AB60961147' x 1142 .
                                 ')|(' . '0201AB60961147' x 10000 . ')');
   %expectpar = (varbincol    => '0x' . '47119660AB0102' x 1142 . '47119660AB01');
   %test      = (varbincol    => '%s =~ /^%s$/');
}
do_tests($X, 1, 'varbinmax', '0x');

%tbl       = (varbincol    => '');
%expectcol = (varbincol    => '0x');
%expectpar = (varbincol    => '0x');
%test      = (varbincol    => '%s eq %s');
do_tests($X, 1, 'varbinmax', 'empty 0x');

$X->{BinaryAsStr} = 0;
%tbl       = (varbincol    => '47119660AB0102' x 10000);
if ($X->{Provider} == PROVIDER_SQLNCLI) {
   %expectcol = (varbincol    => '2010BA06691174' x 10000);
   %expectpar = (varbincol    => '47119660AB0102' x 20000);
   %test      = (varbincol    => '%s eq %s');
}
else {
   %expectcol = (varbincol    => '(' . '691174' . '2010BA06691174' x 571 .
                                 ')|(' . '2010BA06691174' x 10000 . ')');
   %expectpar = (varbincol    => '47119660AB0102' x 571 . '471196');
   %test      = (varbincol    => '%s =~ /^%s$/');
}
do_tests($X, 1, 'varbinmax', 'binary');

%tbl       = (varbincol    => '');
%expectcol = (varbincol    => '');
%expectpar = (varbincol    => '');
%test      = (varbincol    => '%s eq %s');
do_tests($X, 1, 'varbinmax', 'empty bin');


%tbl       = (varbincol    => undef);
%expectcol = (varbincol    => undef);
%expectpar = (varbincol    => undef);
%test      = (varbincol    => 'not defined %s');
do_tests($X, 1, 'varbinmax', 'null');

drop_test_objects('varbinmax');

#------------------------------- UDT -----------------------------------
# We cannot do UDT tests, if the CLR is not enabled on the server.
my $clr_enabled = sql_one(<<SQLEND, Win32::SqlServer::SCALAR);
SELECT value
FROM   sys.configurations
WHERE  name = 'clr enabled'
SQLEND
# At this point we must turn on ANSI_WARNINGS, to get the XML stuff to
# work.
$X->sql("SET ANSI_WARNINGS ON");

goto no_udt if not $clr_enabled;

clear_test_data;
create_UDT1($X, $X->{Provider} == PROVIDER_SQLNCLI);

$X->{BinaryAsStr} = 'x';
%tbl       = (cmplxcol  => '0x800000058000000700',
              pointcol  => '0x01800000098000000480000005',
              stringcol => '0x00050000004E69737365',
              xmlcol    => '<root>input</root>');
%expectcol = (cmplxcol  => '0x800000078000000500',
              pointcol  => '0x0180000012800000088000000A',
              stringcol => '0x0005000000657373694E',
              xmlcol    => '<root>input trigger text</root>');
if ($X->{Provider} == PROVIDER_SQLNCLI) {
   %expectpar = (cmplxcol  => '0x8000000A8000000E00',
                 pointcol  => '0x01800000048000000580000009',
                 stringcol => '0x00050000004E49535345',
                 xmlcol    => '<root>input procedure text</root>');
}
else {
   %expectpar = ();
}
%test      = (cmplxcol  => '%s eq %s',
              pointcol  => '%s eq %s',
              stringcol => '%s eq %s',
              xmlcol    => '%s eq %s');
do_tests($X, 1, 'UDT1', 'Bin 0x');

$X->{BinaryAsStr} = 1;
%tbl       = (cmplxcol  => '0x800000058000000700',
              pointcol  => '0x01800000098000000480000005',
              stringcol => '00A00F0000' . '4E69737365' x 800,
              xmlcol    => '<root>input</root>');
%expectcol = (cmplxcol  => '800000078000000500',
              pointcol  => '0180000012800000088000000A',
              stringcol => '00A00F0000' . '657373694E' x 800,
              xmlcol    => '<root>input trigger text</root>');
if ($X->{Provider} == PROVIDER_SQLNCLI) {
   %expectpar = (cmplxcol  => '8000000A8000000E00',
                 pointcol  => '01800000048000000580000009',
                 stringcol => '00A00F0000' . '4E49535345' x 800,
                 xmlcol    => '<root>input procedure text</root>');
}
else {
   %expectpar = ();
}
%test      = (cmplxcol  => '%s eq %s',
              pointcol  => '%s eq %s',
              stringcol => '%s eq %s',
              xmlcol    => '%s eq %s');
do_tests($X, 1, 'UDT1', 'BinAsStr');

$X->{BinaryAsStr} = 0;
%tbl       = (cmplxcol  => pack('H*', '800000058000000700'),
              pointcol  => pack('H*', '01800000098000000480000005'),
              stringcol => pack('H*', '00050000004E69737365'),
              xmlcol    => '<root>input</root>');
%expectcol = (cmplxcol  => pack('H*', '800000078000000500'),
              pointcol  => pack('H*', '0180000012800000088000000A'),
              stringcol => pack('H*', '0005000000657373694E'),
              xmlcol    => '<root>input trigger text</root>');
if ($X->{Provider} == PROVIDER_SQLNCLI) {
   %expectpar = (cmplxcol  => pack('H*', '8000000A8000000E00'),
                 pointcol  => pack('H*', '01800000048000000580000009'),
                 stringcol => pack('H*', '00050000004E49535345'),
                 xmlcol    => '<root>input procedure text</root>');
}
else {
   %expectpar = ();
}
%test      = (cmplxcol  => '%s eq %s',
              pointcol  => '%s eq %s',
              stringcol => '%s eq %s',
              xmlcol    => '%s eq %s');
do_tests($X, 1, 'UDT1', 'BinaryAsBinary');


%tbl       = (cmplxcol  => undef,
              pointcol  => undef,
              stringcol => undef,
              xmlcol    => undef);
%expectcol = (cmplxcol  => undef,
              pointcol  => undef,
              stringcol => undef,
              xmlcol    => undef);
if ($X->{Provider} == PROVIDER_SQLNCLI) {
   %expectpar = (cmplxcol  => undef,
                 pointcol  => undef,
                 stringcol => undef,
                 xmlcol    => undef);
}
else {
   %expectpar = ();
}
%test      = (cmplxcol  => 'not defined %s',
              pointcol  => 'not defined %s',
              stringcol => 'not defined %s',
              xmlcol    => 'not defined %s');
do_tests($X, 1, 'UDT1', '0x, NULL');


clear_test_data;
create_UDT2($X, $X->{Provider} == PROVIDER_SQLNCLI);

$X->{BinaryAsStr} = 'x';
%tbl       = (cmplxcol  => '0x800000058000000700',
              intcol    => 15,
              stringcol => '0x0000000000');
%expectcol = (cmplxcol  => '0x800000078000000500',
              intcol    => 30,
              stringcol => '0x0000000000');
if ($X->{Provider} == PROVIDER_SQLNCLI) {
   %expectpar = (cmplxcol  => '0x8000000A8000000E00',
                 intcol    => 106,
                 stringcol => '0x0000000000');
}
else {
   %expectpar = ();
}
%test      = (cmplxcol  => '%s eq %s',
              intcol    => '%s eq %s',
              stringcol => '%s eq %s',
              xmlcol    => '%s eq %s');
do_tests($X, 1, 'UDT2', 'Bin0x');

clear_test_data;
create_UDT3($X, $X->{Provider} == PROVIDER_SQLNCLI);
$X->{BinaryAsStr} = 1;
%tbl       = (xmlcol    => '<TEST>Lantliv</TEST>',
              pointcol  => '0x01800000098000000480000005',
              nollcol   => 0);
%expectcol = (xmlcol    => '<TEST>Lantliv trigger text</TEST>',
              pointcol  => '0180000012800000088000000A',
              nollcol   => 19);
if ($X->{Provider} == PROVIDER_SQLNCLI) {
   %expectpar = (xmlcol    => '<TEST>Lantliv procedure text</TEST>',
                 pointcol  => '01800000048000000580000009',
                 nollcol   => -9);
}
else {
   %expectpar = ();
}
%test      = (xmlcol    => '%s eq %s',
              pointcol  => '%s eq %s',
              nollcol   => 'abs(%s - %s) < 1E-9');
do_tests($X, 1, 'UDT3', 'Bin0x');

drop_test_objects('UDT1');
drop_test_objects('UDT2');
drop_test_objects('UDT3');
    sql(<<SQLEND);
    IF EXISTS (SELECT * FROM sys.xml_schema_collections WHERE name = 'OlleSC')
            DROP XML SCHEMA COLLECTION OlleSC
SQLEND
delete_the_udts($X);

no_udt:

#------------------------------- XML -----------------------------------
binmode(STDOUT, ':utf8:');
binmode(STDERR, ':utf8:');

clear_test_data;
create_xmltest($X, $X->{Provider} == PROVIDER_SQLNCLI);

%tbl       = (xmlcol    => "<R\x{00C4}KSM\x{00D6}RG\x{00C5}S>" .
                           "21 pa\x{017A}dziernika 2004 " x 2000 .
                           "</R\x{00C4}KSM\x{00D6}RG\x{00C5}S>",
              xmlsccol  => '<?xml version="1.0" encoding="iso-8859-1"?>' . "\n" .
                            "<TÄST>" .
                            "Vi är alltid bäst i räksmörgåstäster! " x 1500 .
                            "</TÄST>\n<TÄST>I alla fall nästan alltid!</TÄST>",
              nvarcol   => "21 PA\x{0179}DZIERNIKA 2004 " x 2000,
              nvarsccol => "The naïve rôles coöperate with their résumés "
                           x 1000);
%expectcol = (xmlcol    => '<xmltest nvarcol\s*=\s*"' .
                           "21 PA\x{0179}DZIERNIKA 2004 " x 2000 . '"\s*/\s*>',
              xmlsccol  => '<TÄST>' .
                           "The naïve rôles coöperate with their résumés "
                           x 1000 . '</TÄST>',
              nvarcol   => $tbl{'xmlcol'},
              nvarsccol => "Vi är alltid bäst i räksmörgåstäster! " x 1500);
if ($X->{Provider} == PROVIDER_SQLNCLI) {
   %expectpar = (xmlcol   => '<row><Lågland>' .
                             "21 pa\x{017A}dziernika 2004 " x 2000 .
                             '</Lågland></row>',
                 xmlsccol => '<TÄST>' .
                             "THE NAÏVE RÔLES COÖPERATE WITH THEIR RÉSUMÉS "
                             x 1000 . '</TÄST>',
                 nvarcol  => "21 pa\x{017A}dziernika 2004 " x 2000 ,
                 nvarsccol=> "<TÄST>" .
                             "Vi är alltid bäst i räksmörgåstäster! " x 1500 .
                             "</TÄST><TÄST>I alla fall nästan alltid!</TÄST>");
}
else {
   %expectpar = ();
}
%test      = (xmlcol    => '%s =~ %s',
              xmlsccol  => '%s eq %s',
              nvarcol   => '%s eq %s',
              nvarsccol => '%s eq %s');
do_tests($X, 1, 'xmltest');


%tbl       = (xmlcol    => qq!<?xml version = "1.0"\tencoding =   "ucs-2"?>! .
                           "<R\x{00C4}KSM\x{00D6}RG\x{00C5}S>" .
                           "21 pa\x{017A}dziernika 2004 " .
                           "</R\x{00C4}KSM\x{00D6}RG\x{00C5}S>  ",
              xmlsccol  => '<?xml  version="1.0" encoding="UTF-8" ?>' . "\n" .
                            "<TÄST>" .
                            "Vi är alltid bäst i räksmörgåstäster! " .
                            "</TÄST>\n<TÄST>I alla fall nästan alltid!</TÄST>",
              nvarcol   => "   ",
              nvarsccol => 'undef');
%expectcol = (xmlcol    => '<xmltest nvarcol\s*=\s*"   "\s*/\s*>',
              xmlsccol  => '<TÄST>undef</TÄST>',
              nvarcol   => "<R\x{00C4}KSM\x{00D6}RG\x{00C5}S>" .
                           "21 pa\x{017A}dziernika 2004 " .
                           "</R\x{00C4}KSM\x{00D6}RG\x{00C5}S>",
              nvarsccol => "Vi är alltid bäst i räksmörgåstäster! ");
if ($X->{Provider} == PROVIDER_SQLNCLI) {
   %expectpar = (xmlcol   => '<row><Lågland>( |\&\#x20;){3,3}</Lågland></row>',
                 xmlsccol => '<TÄST>UNDEF</TÄST>',
                 nvarcol  => "21 pa\x{017A}dziernika 2004 ",
                 nvarsccol=> "<TÄST>" .
                             "Vi är alltid bäst i räksmörgåstäster! ".
                             "</TÄST><TÄST>I alla fall nästan alltid!</TÄST>");
}
else {
   %expectpar = ();
}
%test      = (xmlcol    => '%s =~ %s',
              xmlsccol  => '%s eq %s',
              nvarcol   => '%s eq %s',
              nvarsccol => '%s eq %s');
do_tests($X, 1, 'xmltest', 'take two');


%tbl       = (xmlcol    => '',
              xmlsccol  => '   ',
              nvarcol   => undef,
              nvarsccol => undef);
%expectcol = (xmlcol    => '<xmltest\s*/\s*>',
              xmlsccol  => '<TÄST/>',
              nvarcol   => undef,
              nvarsccol => undef);
if ($X->{Provider} == PROVIDER_SQLNCLI) {
   %expectpar = (xmlcol   => '<row\s*/\s*>',
                 xmlsccol => '<TÄST/>',
                 nvarcol  => undef,
                 nvarsccol=> undef);
}
else {
   %expectpar = ();
}
%test      = (xmlcol    => '%s =~ %s',
              xmlsccol  => '%s eq %s',
              nvarcol   => 'not defined %s',
              nvarsccol => 'not defined %s');
do_tests($X, 1, 'xmltest', 'empty strings');


drop_test_objects('xmltest');
    sql(<<SQLEND);
    IF EXISTS (SELECT * FROM sys.xml_schema_collections WHERE name = 'Olles SC')
            DROP XML SCHEMA COLLECTION [Olles SC]
SQLEND

#-----------------------------------------------------------------------
#-----------------------------------------------------------------------
# Finally test parameterless SP.
finally:
{
   blurb("parameterless SP");
   sql ("CREATE PROCEDURE #pelle_sp AS SELECT 4711");
   my $result;
   $result = sql_sp('#pelle_sp', SCALAR, SINGLEROW);
   push(@testres, ($result == 4711 ? "ok %d" : "not ok %d"));
   $no_of_tests++;
}


print "1..$no_of_tests\n";

my $no = 1;
foreach my $line (@testres) {
   printf "$line\n", $no;
   $no++ if $line =~ /^(not )?ok/;
}

