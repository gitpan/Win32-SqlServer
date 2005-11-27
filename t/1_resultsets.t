#---------------------------------------------------------------------
# $Header: /Perl/OlleDB/t/1_resultsets.t 7     05-11-26 23:47 Sommar $
#
# $History: 1_resultsets.t $
# 
# *****************  Version 7  *****************
# User: Sommar       Date: 05-11-26   Time: 23:47
# Updated in $/Perl/OlleDB/t
# Renamed the module from MSSQL::OlleDB to Win32::SqlServer.
#
# *****************  Version 6  *****************
# User: Sommar       Date: 05-08-06   Time: 23:23
# Updated in $/Perl/OlleDB/t
# Added test for sql_sp and callback.
#
# *****************  Version 5  *****************
# User: Sommar       Date: 05-02-27   Time: 21:54
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 4  *****************
# User: Sommar       Date: 05-02-06   Time: 20:45
# Updated in $/Perl/OlleDB/t
#
# *****************  Version 3  *****************
# User: Sommar       Date: 05-01-02   Time: 20:56
# Updated in $/Perl/OlleDB/t
# Small adjustment to the require.
#
# *****************  Version 2  *****************
# User: Sommar       Date: 05-01-02   Time: 20:53
# Updated in $/Perl/OlleDB/t
# Now login is controlled from environment variable.
#
# *****************  Version 1  *****************
# User: Sommar       Date: 05-01-02   Time: 20:27
# Created in $/Perl/OlleDB/t
#---------------------------------------------------------------------

use strict;
use Win32::SqlServer qw(:DEFAULT :consts);
use File::Basename qw(dirname);

require &dirname($0) . '\testsqllogin.pl';

use vars qw(@testres $verbose);

sub blurb{
    push (@testres, "#------ Testing @_ ------\n");
    print "#------ Testing @_ ------\n" if $verbose;
}

$verbose = shift @ARGV;

$^W = 1;

$| = 1;

my($X, $sql, $sql1, $sql_empty, $sql_error, $sql_null, $sql_key1, $sql_key_many);

$X = testsqllogin();

# Accept all errors, and be silent about them.
$X->{errInfo}{maxSeverity}   = 25;
$X->{errInfo}{printLines} = 25;
$X->{errInfo}{printMsg}   = 25;
$X->{errInfo}{printText}  = 25;
$X->{errInfo}{carpLevel}  = 25;

$SQLSEP = '@!@';

# First set up tables and data.
sql(<<SQLEND);
CREATE TABLE #a(a char(1), b char(1), i int)
CREATE TABLE #b(x char(3) NULL)
CREATE TABLE #c(key1  char(5)     NOT NULL,
                key2  char(1)     NOT NULL,
                key3  int         NOT NULL,
                data1 smallint    NULL,
                data2 varchar(10) NULL,
                data3 char(1)     NOT NULL)

INSERT #a VALUES('A', 'A', 12)
INSERT #a VALUES('A', 'D', 24)
INSERT #a VALUES('A', 'H', 1)
INSERT #a VALUES('C', 'B', 12)

INSERT #b VALUES('xyz')
INSERT #b VALUES(NULL)

INSERT #c VALUES('apple', 'X', 1, NULL, NULL,      'T')
INSERT #c VALUES('apple', 'X', 2, -15,  NULL,      'T')
INSERT #c VALUES('apple', 'X', 3, NULL, NULL,      'T')
INSERT #c VALUES('apple', 'Y', 1, 18,   'Verdict', 'H')
INSERT #c VALUES('apple', 'Y', 6, 18,   'Maracas', 'I')
INSERT #c VALUES('peach', 'X', 1, 18,   'Lastkey', 'T')
INSERT #c VALUES('peach', 'X', 8, 4711, 'Monday',  'T')
INSERT #c VALUES('melon', 'Y', 1, 118,  'Lastkey', 'T')
SQLEND

# This is our test batch: three result sets whereof one empty.
$sql = <<SQLEND;
SELECT *
FROM   #a
ORDER  BY a, b
COMPUTE SUM(i) BY a
COMPUTE SUM(i)

SELECT * FROM #b

-- Note: if this SELECT comes directly after the first SELECT, SQLOLEDB
-- gets an AV. Not our fault. :-)
SELECT * FROM #a WHERE a = '?'

SELECT 4711
SQLEND

# Test code for single-row queries.
$sql1 = "SELECT * FROM #a WHERE i = 24";

# Test for SELECT of NULL.
$sql_null = "SELECT NULL";

# Test code for empty result sets
$sql_empty = <<SQLEND;
SELECT * FROM #a WHERE i = 456
SELECT * FROM #a WHERE a = 'z'
SQLEND

# Test code with incorrect SQL which will not produce even a resultset,
$sql_error = 'SELECT FROM';

# Test code for keyed access.
$sql_key1     = "SELECT * FROM #a";
sql("CREATE PROCEDURE #sql_key_many AS SELECT * FROM #c");

#-------------------- MULTISET ----------------------------
{
   my (@result, $result, @expect);

   &blurb("HASH, MULTISET, wantarray");
   @expect = ([{a => 'A', b => 'A', i => 12},
               {a => 'A', b => 'D', i => 24},
               {a => 'A', b => 'H', i => 1}],
              [{sum => 37}],
              [{a => 'C', b => 'B', i => 12}],
              [{sum => 12}],
              [{sum => 49}],
              [{x => 'xyz'},
               {x => undef}],
              [],
              [{'Col 1' => 4711}]);
   @result = sql($sql, HASH, MULTISET);
   push(@testres, compare(\@expect, \@result));

   &blurb("HASH, MULTISET, wantscalar");
   $result = sql($sql, HASH, MULTISET);
   push(@testres, compare(\@expect, $result));


   &blurb("LIST, MULTISET, wantarray");
   @expect = ([['A', 'A', 12],
               ['A', 'D', 24],
               ['A', 'H', 1]],
              [[37]],
              [['C', 'B', 12]],
              [[12]],
              [[49]],
              [['xyz'],
               [undef]],
              [],
              [[4711]]);
   @result = sql($sql, LIST, MULTISET);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, MULTISET, wantscalar");
   $result = sql($sql, LIST, MULTISET);
   push(@testres, compare(\@expect, $result));

   &blurb("SCALAR, MULTISET, wantarray");
   @expect = (['A@!@A@!@12',
               'A@!@D@!@24',
               'A@!@H@!@1'],
              ['37'],
              ['C@!@B@!@12'],
              ['12'],
              ['49'],
              ['xyz',
               undef],
              [],
              ['4711']);
   @result = sql($sql, MULTISET, SCALAR);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, MULTISET, wantscalar");
   $result = sql($sql, SCALAR, MULTISET);
   push(@testres, compare(\@expect, $result));
}

#--------------------- MULTISET empty, empty ------------------------
{
   my (@result, $result, @expect);

   @expect = ([], []);
   &blurb("HASH, MULTISET empty, wantarray");
   @result = sql($sql_empty, HASH, MULTISET);
   push(@testres, compare(\@expect, \@result));

   &blurb("HASH, MULTISET empty, wantscalar");
   $result = sql($sql_empty, HASH, MULTISET);
   push(@testres, compare(\@expect, $result));

   &blurb("LIST, MULTISET empty, wantarray");
   @result = sql($sql_empty, LIST, MULTISET);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, MULTISET empty, wantscalar");
   $result = sql($sql_empty, LIST, MULTISET);
   push(@testres, compare(\@expect, $result));

   &blurb("SCALAR, MULTISET empty, wantarray");
   @result = sql($sql_empty, SCALAR, MULTISET);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, MULTISET empty, wantscalar");
   $result = sql($sql_empty, SCALAR, MULTISET);
   push(@testres, compare(\@expect, $result));
}

#--------------------- MULTISET error   ------------------------
{
   my (@result, $result, @expect);

   @expect = ([]);
   &blurb("HASH, MULTISET error, wantarray");
   @result = sql($sql_error, HASH, MULTISET);
   push(@testres, compare(\@expect, \@result));

   &blurb("HASH, MULTISET error, wantscalar");
   $result = sql($sql_error, HASH, MULTISET);
   push(@testres, compare(\@expect, $result));

   &blurb("LIST, MULTISET error, wantarray");
   @result = sql($sql_error, LIST, MULTISET);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, MULTISET error, wantscalar");
   $result = sql($sql_error, LIST, MULTISET);
   push(@testres, compare(\@expect, $result));

   &blurb("SCALAR, MULTISET error, wantarray");
   @result = sql($sql_error, SCALAR, MULTISET);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, MULTISET error, wantscalar");
   $result = sql($sql_error, SCALAR, MULTISET);
   push(@testres, compare(\@expect, $result));
}

#--------------------- MULTISET noexec   ------------------------
{
   my (@result, $result, @expect);

   $X->{NoExec} = 1;
   @expect = ();
   &blurb("HASH, MULTISET NoExec, wantarray");
   @result = sql($sql, HASH, MULTISET);
   push(@testres, compare(\@expect, \@result));

   &blurb("HASH, MULTISET NoExec, wantscalar");
   $result = sql($sql, HASH, MULTISET);
   push(@testres, compare(undef, $result));

   &blurb("LIST, MULTISET NoExec, wantarray");
   @result = sql($sql, LIST, MULTISET);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, MULTISET NoExec, wantscalar");
   $result = sql($sql, LIST, MULTISET);
   push(@testres, compare(undef, $result));

   &blurb("SCALAR, MULTISET NoExec, wantarray");
   @result = sql($sql, SCALAR, MULTISET);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, MULTISET NoExec, wantscalar");
   $result = sql($sql, SCALAR, MULTISET);
   push(@testres, compare(undef, $result));
   $X->{NoExec} = 0;
}

#-------------------- SINGLESET ----------------------------
{
   my (@result, $result, @expect);

   &blurb("HASH, SINGLESET, wantarray");
   @expect = ({a => 'A', b => 'A', i => 12},
              {a => 'A', b => 'D', i => 24},
              {a => 'A', b => 'H', i => 1},
              {sum => 37},
              {a => 'C', b => 'B', i => 12},
              {sum => 12},
              {sum => 49},
              {'x' => 'xyz'},
              {'x' => undef},
              {'Col 1' => 4711});
   @result = sql($sql);
   push(@testres, compare(\@expect, \@result));

   &blurb("HASH, SINGLESET, wantscalar");
   $result = sql($sql);
   push(@testres, compare(\@expect, $result));

   &blurb("LIST, SINGLESET, wantarray");
   @expect = (['A', 'A', 12],
              ['A', 'D', 24],
              ['A', 'H', 1],
              [37],
              ['C', 'B', 12],
              [12],
              [49],
              ['xyz'],
              [undef],
              [4711]);
   @result = sql($sql, LIST);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, SINGLESET, wantscalar");
   $result = sql($sql, undef, LIST);
   push(@testres, compare(\@expect, $result));

   &blurb("SCALAR, SINGLESET, wantarray");
   @expect = ('A@!@A@!@12',
              'A@!@D@!@24',
              'A@!@H@!@1',
              '37',
              'C@!@B@!@12',
              '12',
              '49',
              'xyz',
              undef,
              '4711');
   @result = sql($sql, SCALAR);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, SINGLESET, wantscalar");
   $result = sql($sql, SCALAR);
   push(@testres, compare(\@expect, $result));
}

#--------------------- SINGLESET, empty ------------------------
{
   my (@result, $result, @expect);

   @expect = ();
   &blurb("HASH, SINGLESET empty, wantarray");
   @result = sql($sql_empty, HASH, SINGLESET);
   push(@testres, compare(\@expect, \@result));

   &blurb("HASH, SINGLESET empty, wantscalar");
   $result = sql($sql_empty, HASH, SINGLESET);
   push(@testres, compare(\@expect, $result));

   &blurb("LIST, SINGLESET empty, wantarray");
   @result = sql($sql_empty, LIST, SINGLESET);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, SINGLESET empty, wantscalar");
   $result = sql($sql_empty, LIST, SINGLESET);
   push(@testres, compare(\@expect, $result));

   &blurb("SCALAR, SINGLESET empty, wantarray");
   @result = sql($sql_empty, SCALAR, SINGLESET);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, SINGLESET empty, wantscalar");
   $result = sql($sql_empty, SCALAR, SINGLESET);
   push(@testres, compare(\@expect, $result));
}

#-------------------- SINGLESET, error ----------------------
{
   my (@result, $result, @expect);

   @expect = ();
   &blurb("HASH, SINGLESET error, wantarray");
   @result = sql($sql_error, HASH, SINGLESET);
   push(@testres, compare(\@expect, \@result));

   &blurb("HASH, SINGLESET error, wantscalar");
   $result = sql($sql_error, HASH, SINGLESET);
   push(@testres, compare(\@expect, $result));

   &blurb("LIST, SINGLESET error, wantarray");
   @result = sql($sql_error, LIST, SINGLESET);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, SINGLESET error, wantscalar");
   $result = sql($sql_error, LIST, SINGLESET);
   push(@testres, compare(\@expect, $result));

   &blurb("SCALAR, SINGLESET error, wantarray");
   @result = sql($sql_error, SCALAR, SINGLESET);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, SINGLESET error, wantscalar");
   $result = sql($sql_error, SCALAR, SINGLESET);
   push(@testres, compare(\@expect, $result));
}

#-------------------- SINGLESET, NoExec ----------------------
{
   my (@result, $result, @expect);

   $X->{NoExec} = 1;
   @expect = ();
   &blurb("HASH, SINGLESET NoExec, wantarray");
   @result = sql($sql, HASH, SINGLESET);
   push(@testres, compare(\@expect, \@result));

   &blurb("HASH, SINGLESET NoExec, wantscalar");
   $result = sql($sql, HASH, SINGLESET);
   push(@testres, compare(undef, $result));

   &blurb("LIST, SINGLESET NoExec, wantarray");
   @result = sql($sql, LIST, SINGLESET);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, SINGLESET NoExec, wantscalar");
   $result = sql($sql, LIST, SINGLESET);
   push(@testres, compare(undef, $result));

   &blurb("SCALAR, SINGLESET NoExec, wantarray");
   @result = sql($sql, SCALAR, SINGLESET);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, SINGLESET NoExec, wantscalar");
   $result = sql($sql, SCALAR, SINGLESET);
   push(@testres, compare(undef, $result));
   $X->{NoExec} = 0;
}

#-------------------- SINGLEROW ----------------------------
{
   my (@result, %result, $result, @expect, %expect, $expect);

   &blurb("HASH, SINGLEROW, wantarray");
   %expect = (a => 'A', b => 'D', i => 24);
   %result = sql($sql1, undef, SINGLEROW);
   push(@testres, compare(\%expect, \%result));

   &blurb("HASH, SINGLEROW, wantscalar");
   $result = sql($sql1, SINGLEROW, undef);
   push(@testres, compare(\%expect, $result));

   &blurb("LIST, SINGLEROW, wantarray");
   @expect = ('A', 'D', 24);
   @result = sql($sql1, LIST, SINGLEROW);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, SINGLEROW, wantscalar");
   $result = sql($sql1, LIST, SINGLEROW);
   push(@testres, compare(\@expect, $result));

   &blurb("SCALAR, SINGLEROW, wantarray");
   @expect = ('A@!@D@!@24');
   @result = sql($sql1, SCALAR, SINGLEROW);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, SINGLEROW, wantscalar");
   $expect = 'A@!@D@!@24';
   $result = sql($sql1, SCALAR, SINGLEROW);
   push(@testres, compare($expect, $result));
}

#-------------------- SINGLEROW, SELECT NULL---------------------------
{
   my (@result, %result, $result, @expect, %expect, $expect);

   &blurb("HASH, SINGLEROW NULL, wantarray");
   %expect = ('Col 1' => undef);
   %result = sql($sql_null, undef, SINGLEROW);
   push(@testres, compare(\%expect, \%result));

   &blurb("HASH, SINGLEROW NULL, wantscalar");
   $result = sql($sql_null, SINGLEROW, undef);
   push(@testres, compare(\%expect, $result));

   &blurb("LIST, SINGLEROW NULL, wantarray");
   @expect = (undef);
   @result = sql($sql_null, LIST, SINGLEROW);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, SINGLEROW NULL, wantscalar");
   $result = sql($sql_null, LIST, SINGLEROW);
   push(@testres, compare(\@expect, $result));

   &blurb("SCALAR, SINGLEROW NULL, wantarray");
   @expect = (undef);
   @result = sql($sql_null, SCALAR, SINGLEROW);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, SINGLEROW NULL, wantscalar");
   $expect = undef;
   $result = sql($sql_null, SCALAR, SINGLEROW);
   push(@testres, compare($expect, $result));
}

#--------------- SINGLEROW, first result empty ----------------------------
{
   my (@result, %result, $result, @expect, %expect, $expect);

   &blurb("HASH, SINGLEROW first empty, wantarray");
   %expect = (a => 'A', b => 'D', i => 24);
   %result = sql("$sql_empty $sql1", undef, SINGLEROW);
   push(@testres, compare(\%expect, \%result));

   &blurb("HASH, SINGLEROW first empty, wantscalar");
   $result = sql("$sql_empty $sql1", SINGLEROW, undef);
   push(@testres, compare(\%expect, $result));

   &blurb("LIST, SINGLEROW first empty, wantarray");
   @expect = ('A', 'D', 24);
   @result = sql("$sql_empty $sql1", LIST, SINGLEROW);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, SINGLEROW first empty, wantscalar");
   $result = sql("$sql_empty $sql1", LIST, SINGLEROW);
   push(@testres, compare(\@expect, $result));

   &blurb("SCALAR, SINGLEROW first empty, wantarray");
   @expect = ('A@!@D@!@24');
   @result = sql("$sql_empty $sql1", SCALAR, SINGLEROW);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, SINGLEROW first empty, wantscalar");
   $expect = 'A@!@D@!@24';
   $result = sql("$sql_empty $sql1", SCALAR, SINGLEROW);
   push(@testres, compare($expect, $result));
}

#--------------------- SINGLEROW, empty ------------------------
{
   my (@result, %result, $result, @expect, %expect);
   @expect = %expect = ();

   &blurb("HASH, SINGLEROW empty, wantarray");
   %result = sql($sql_empty, HASH, SINGLEROW);
   push(@testres, compare(\%expect, \%result));

   &blurb("HASH, SINGLEROW empty, wantscalar");
   $result = sql($sql_empty, HASH, SINGLEROW);
   push(@testres, compare(undef, $result));

   &blurb("LIST, SINGLEROW empty, wantarray");
   @result = sql($sql_empty, LIST, SINGLEROW);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, SINGLEROW empty, wantscalar");
   $result = sql($sql_empty, LIST, SINGLEROW);
   push(@testres, compare(undef, $result));

   &blurb("SCALAR, SINGLEROW empty, wantarray");
   @result = sql($sql_empty, SCALAR, SINGLEROW);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, SINGLEROW empty, wantscalar");
   $result = sql($sql_empty, SCALAR, SINGLEROW);
   push(@testres, compare(undef, $result));
}

#--------------------- SINGLEROW, error -------------------
{
   my (@result, %result, $result, @expect, %expect);

   @expect = %expect = ();

   &blurb("HASH, SINGLEROW error, wantarray");
   %result = sql($sql_error, HASH, SINGLEROW);
   push(@testres, compare(\%expect, \%result));

   &blurb("HASH, SINGLEROW error, wantscalar");
   $result = sql($sql_error, HASH, SINGLEROW);
   push(@testres, compare(undef, $result));

   &blurb("LIST, SINGLEROW error, wantarray");
   @result = sql($sql_error, LIST, SINGLEROW);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, SINGLEROW error, wantscalar");
   $result = sql($sql_error, LIST, SINGLEROW);
   push(@testres, compare(undef, $result));

   &blurb("SCALAR, SINGLEROW error, wantarray");
   @result = sql($sql_error, SCALAR, SINGLEROW);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, SINGLEROW error, wantscalar");
   $result = sql($sql_error, SCALAR, SINGLEROW);
   push(@testres, compare(undef, $result));
}

#--------------------- SINGLEROW, NoExec -------------------
{
   my (@result, %result, $result, @expect, %expect);

   $X->{NoExec} = 1;
   @expect = %expect = ();
   &blurb("HASH, SINGLEROW NoExec, wantarray");
   %result = sql($sql, HASH, SINGLEROW);
   push(@testres, compare(\%expect, \%result));

   &blurb("HASH, SINGLEROW NoExec, wantscalar");
   $result = sql($sql, HASH, SINGLEROW);
   push(@testres, compare(undef, $result));

   &blurb("LIST, SINGLEROW NoExec, wantarray");
   @result = sql($sql, LIST, SINGLEROW);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, SINGLEROW NoExec, wantscalar");
   $result = sql($sql, LIST, SINGLEROW);
   push(@testres, compare(undef, $result));

   &blurb("SCALAR, SINGLEROW NoExec, wantarray");
   @result = sql($sql, SCALAR, SINGLEROW);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, SINGLEROW NoExec, wantscalar");
   $result = sql($sql, SCALAR, SINGLEROW);
   push(@testres, compare(undef, $result));
   $X->{NoExec} = 0;
}

#-------------------- sql_one ----------------------------
{
   my (@result, %result, $result, @expect, %expect, $expect);

   &blurb("HASH, sql_one, wantarray");
   %expect = (a => 'A', b => 'D', i => 24);
   %result = sql_one($sql1);
   push(@testres, compare(\%expect, \%result));

   &blurb("HASH, sql_one, wantscalar");
   $result = sql_one($sql1, HASH);
   push(@testres, compare(\%expect, $result));

   &blurb("LIST, sql_one, wantarray");
   @expect = ('A', 'D', 24);
   @result = sql_one($sql1, LIST);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, sql_one, wantscalar");
   $result = sql_one($sql1, LIST);
   push(@testres, compare(\@expect, $result));

   &blurb("SCALAR, sql_one, wantscalar");
   $expect = 'A@!@D@!@24';
   $result = sql_one($sql1);
   push(@testres, compare($expect, $result));

   &blurb("SCALAR, sql_one, two ressets, one row");
   $result = sql_one("SELECT * FROM #b WHERE 1 = 0 $sql1");
   push(@testres, compare($expect, $result));

   &blurb("SCALAR, sql_one, one-NULL col, wantscalar");
   $expect = undef;
   $result = sql_one("SELECT NULL");
   push(@testres, compare($expect, $result));

   &blurb("SCALAR, sql_one, two-NULL cols, wantscalar");
   $expect = '@!@';
   $result = sql_one("SELECT NULL, NULL");
   push(@testres, compare($expect, $result));

   &blurb("sql_one, fail: no rows");
   eval("sql_one('SELECT * FROM #a WHERE i = 897')");
   push(@testres, ($@ =~ /returned no/ ? 1 : 0));

   &blurb("sql_one, fail: too many rows");
   eval("sql_one('SELECT * FROM #a')");
   push(@testres, ($@ =~ /more than one/ ? 1 : 0));

   &blurb("sql_one, fail: two ressets, two rows");
   eval("sql_one('SELECT 1 SELECT 2')");
   push(@testres, ($@ =~ /more than one/ ? 1 : 0));

   &blurb("sql_one, fail: syntax error => no rows");
   eval("sql_one('$sql_error')");
   push(@testres, ($@ =~ /returned no/ ? 1 : 0));

   &blurb("sql_one, fail: type error => no rwows.");
   eval("sql_one('SELECT * FROM #a WHERE i = ?', [['notype', 2]])");
   push(@testres, ($@ =~ /returned no/ ? 1 : 0));
}

#-------------------- sql_one NoExec, noexec ----------------------------
{
   my (@result, %result, $result, @expect, %expect, $expect);
   $X->{NoExec} = 1;
   @expect = %expect = ();
   $expect = undef;

   &blurb("HASH, sql_one NoExec, wantarray");
   %result = sql_one($sql1);
   push(@testres, compare(\%expect, \%result));

   &blurb("HASH, sql_one NoExec, wantscalar");
   $result = sql_one($sql1, HASH);
   push(@testres, compare(undef, $result));

   &blurb("LIST, sql_one NoExec, wantarray");
   @result = sql_one($sql1, LIST);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, sql_one NoExec, wantscalar");
   $result = sql_one($sql1, LIST);
   push(@testres, compare(undef, $result));

   &blurb("SCALAR, sql_one NoExec, wantscalar");
   $result = sql_one($sql1);
   push(@testres, compare($expect, $result));

   &blurb("SCALAR, sql_one NoExec, two ressets, one row");
   $result = sql_one("SELECT * FROM #b WHERE 1 = 0 $sql1");
   push(@testres, compare($expect, $result));

   &blurb("sql_one NoExec, no rows");
   $result = sql_one('SELECT * FROM #a WHERE i = 897');
   push(@testres, compare($expect, $result));

   &blurb("sql_one NoExec, too many rows");
   $result = sql_one('SELECT * FROM #a');
   push(@testres, compare($expect, $result));

   &blurb("sql_one NoExec: two ressets, two rows");
   %result = sql_one('SELECT 1 SELECT 2');
   push(@testres, compare(\%expect, \%result));

   &blurb("sql_one NoExec, fail: syntax error => no rows");
   $result = sql_one('$sql_error');
   push(@testres, compare($expect, $result));

   &blurb("sql_one NoExec, type error => no rwows.");
   $result = sql_one('SELECT * FROM #a WHERE i = ?', [['notype', 2]]);
   push(@testres, compare($expect, $result));
   $X->{NoExec} = 0;
}

#-------------------- NORESULT ----------------------------
{
   my (@result, %result, $result, @expect, %expect, $expect);

   &blurb("HASH, NORESULT, wantarray");
   @expect = %expect = ();
   $expect = undef;
   %result = sql($sql, HASH, NORESULT);
   push(@testres, compare(\%expect, \%result));

   &blurb("HASH, NORESULT, wantscalar");
   $result = sql($sql, HASH, NORESULT);
   push(@testres, compare($expect, $result));

   &blurb("LIST, NORESULT, wantarray");
   @result = sql($sql, LIST, NORESULT);
   push(@testres, compare(\@expect, \@result));

   &blurb("LIST, NORESULT, wantscalar");
   $result = sql($sql, LIST, NORESULT);
   push(@testres, compare($expect, $result));

   &blurb("SCALAR, NORESULT, wantarray");
   @result = sql($sql, SCALAR, NORESULT);
   push(@testres, compare(\@expect, \@result));

   &blurb("SCALAR, NORESULT, wantscalar");
   $result = sql($sql, SCALAR, NORESULT);
   push(@testres, compare($expect, $result));
}

#---------------------- KEYED, single key -------------------
{
   my (%result, $result, %expect);

   &blurb("HASH, KEYED, single key, wantarray");
   %expect = ('A' => {'a' => 'A', 'i' => 12},
              'D' => {'a' => 'A', 'i' => 24},
              'H' => {'a' => 'A', 'i' => 1},
              'B' => {'a' => 'C', 'i' => 12});
   %result = sql($sql_key1, KEYED, ['b']);
   push(@testres, compare(\%expect, \%result));

   &blurb("HASH, KEYED, single key, wantref");
   $result = sql($sql_key1, HASH, KEYED, ['b']);
   push(@testres, compare(\%expect, $result));

   &blurb("LIST, KEYED, single key, wantarray");
   %expect = ('A' => ['A', 12],
              'D' => ['A', 24],
              'H' => ['A', 1],
              'B' => ['C', 12]);
   %result = sql($sql_key1, LIST, KEYED, [2]);
   push(@testres, compare(\%expect, \%result));

   &blurb("LIST, KEYED, single key, wantref");
   $result = sql($sql_key1, LIST, KEYED, [2]);
   push(@testres, compare(\%expect, $result));

   &blurb("SCALAR, KEYED, single key, wantarray");
   %expect = ('A' => 'A@!@12',
              'D' => 'A@!@24',
              'H' => 'A@!@1',
              'B' => 'C@!@12');
   %result = sql($sql_key1, SCALAR, KEYED, [2]);
   push(@testres, compare(\%expect, \%result));

   &blurb("SCALAR, KEYED, single key, wantref");
   $result = sql($sql_key1, SCALAR, KEYED, [2]);
   push(@testres, compare(\%expect, $result));
}

#---------------------- KEYED, multiple key -------------------
{
   my (%result, $result, %expect);

   &blurb("HASH, KEYED, multiple key, wantarray");
   %expect = ('apple' => {'X' => {'1' => {data1 => undef, data2 => undef,     data3 => 'T'},
                                  '2' => {data1 => -15,   data2 => undef,     data3 => 'T'},
                                  '3' => {data1 => undef, data2 => undef,     data3 => 'T'}
                                 },
                          'Y' => {'1' => {data1 => 18,    data2 => 'Verdict', data3 => 'H'},
                                  '6' => {data1 => 18,    data2 => 'Maracas', data3 => 'I'}
                                 }
                         },
              'peach' => {'X' => {'1' => {data1 => 18,    data2 => 'Lastkey', data3 => 'T'},
                                  '8' => {data1 => 4711,  data2 => 'Monday',  data3 => 'T'}
                                  }
                         },
              'melon' => {'Y' => {'1' => {data1 => 118,   data2 => 'Lastkey',  data3 => 'T'}
                                 }
                         }
             );
   %result = sql_sp('#sql_key_many', HASH, KEYED, ['key1', 'key2', 'key3']);
   push(@testres, compare(\%expect, \%result));

   &blurb("HASH, KEYED, multiple key, wantref");
   $result = sql_sp('#sql_key_many', HASH, KEYED, ['key1', 'key2', 'key3']);
   push(@testres, compare(\%expect, $result));

   &blurb("LIST, KEYED, mulitple key, wantarray");
   %expect = ('apple' => {'X' => {'1' => [undef, undef,    'T'],
                                  '2' => [-15,   undef,    'T'],
                                  '3' => [undef, undef,    'T']
                                 },
                          'Y' => {'1' => [18,   'Verdict', 'H'],
                                  '6' => [18,   'Maracas', 'I']
                                 }
                         },
              'peach' => {'X' => {'1' => [18,   'Lastkey', 'T'],
                                  '8' => [4711, 'Monday',  'T']
                                  }
                         },
              'melon' => {'Y' => {'1' => [118,  'Lastkey', 'T']
                                 }
                         }
             );
   %result = sql_sp('#sql_key_many', LIST, KEYED, [1, 2, 3]);
   push(@testres, compare(\%expect, \%result));

   &blurb("LIST, KEYED, multiple key, wantref");
   $result = sql_sp('#sql_key_many', LIST, KEYED, [1, 2, 3]);
   push(@testres, compare(\%expect, $result));

   &blurb("SCALAR, KEYED, multiple key, wantarray");
   %expect = ('apple' => {'X' => {'1' => '@!@@!@T',
                                  '2' => '-15@!@@!@T',
                                  '3' => '@!@@!@T'
                                 },
                          'Y' => {'1' => '18@!@Verdict@!@H',
                                  '6' => '18@!@Maracas@!@I'
                                 }
                         },
              'peach' => {'X' => {'1' => '18@!@Lastkey@!@T',
                                  '8' => '4711@!@Monday@!@T'
                                  }
                         },
              'melon' => {'Y' => {'1' => '118@!@Lastkey@!@T'
                                 }
                         }
             );
   %result = sql_sp('#sql_key_many', SCALAR, KEYED, [1, 2, 3]);
   push(@testres, compare(\%expect, \%result));

   &blurb("SCALAR, KEYED, multiple key, wantref");
   $result = sql_sp('#sql_key_many', SCALAR, KEYED, [1, 2, 3]);
   push(@testres, compare(\%expect, $result));
}

#-------------------- KEYED, empty ----------------------
{
   my (%result, $result, %expect);

   %expect = ();
   &blurb("HASH, KEYED empty, wantarray");
   %result = sql($sql_empty, HASH, KEYED, ['a']);
   push(@testres, compare(\%expect, \%result));

   &blurb("HASH, KEYED empty, wantscalar");
   $result = sql($sql_empty, HASH, KEYED, ['a']);
   push(@testres, compare(\%expect, $result));

   &blurb("LIST, KEYED empty, wantarray");
   %result = sql($sql_empty, LIST, KEYED, [1]);
   push(@testres, compare(\%expect, \%result));

   &blurb("LIST, KEYED empty, wantscalar");
   $result = sql($sql_empty, LIST, KEYED, [1]);
   push(@testres, compare(\%expect, $result));

   &blurb("SCALAR, KEYED empty, wantarray");
   %result = sql($sql_empty, SCALAR, KEYED, [1]);
   push(@testres, compare(\%expect, \%result));

   &blurb("SCALAR, KEYED empty, wantscalar");
   $result = sql($sql_empty, SCALAR, KEYED, [1]);
   push(@testres, compare(\%expect, $result));
}

#--------------------- KEYED, sql_error  -------------------
{
   my (%result, $result, %expect);

   %expect = ();
   &blurb("HASH, KEYED error, wantarray");
   %result = sql($sql_error, HASH, KEYED, ['a']);
   push(@testres, compare(\%expect, \%result));

   &blurb("HASH, KEYED error, wantscalar");
   $result = sql($sql_error, HASH, KEYED, ['a']);
   push(@testres, compare(\%expect, $result));

   &blurb("LIST, KEYED error, wantarray");
   %result = sql($sql_error, LIST, KEYED, [1]);
   push(@testres, compare(\%expect, \%result));

   &blurb("LIST, KEYED error, wantscalar");
   $result = sql($sql_error, LIST, KEYED, [1]);
   push(@testres, compare(\%expect, $result));

   &blurb("SCALAR, KEYED error, wantarray");
   %result = sql($sql_error, SCALAR, KEYED, [1]);
   push(@testres, compare(\%expect, \%result));

   &blurb("SCALAR, KEYED error, wantscalar");
   $result = sql($sql_error, SCALAR, KEYED, [1]);
   push(@testres, compare(\%expect, $result));
}

#--------------------- KEYED, NoExec -------------------
{
   my (%result, $result, %expect);

   $X->{NoExec} = 1;
   %expect = ();
   &blurb("HASH, KEYED NoExec, wantarray");
   %result = sql($sql_key1, HASH, KEYED, ['a']);
   push(@testres, compare(\%expect, \%result));

   &blurb("HASH, KEYED NoExec, wantscalar");
   $result = sql($sql_key1, HASH, KEYED, ['a']);
   push(@testres, compare(undef, $result));

   &blurb("LIST, KEYED NoExec, wantarray");
   %result = sql($sql_key1, LIST, KEYED, [1]);
   push(@testres, compare(\%expect, \%result));

   &blurb("LIST, KEYED NoExec, wantscalar");
   $result = sql($sql_key1, LIST, KEYED, [1]);
   push(@testres, compare(undef, $result));

   &blurb("SCALAR, KEYED NoExec, wantarray");
   %result = sql($sql_key1, SCALAR, KEYED, [1]);
   push(@testres, compare(\%expect, \%result));

   &blurb("SCALAR, KEYED NoExec, wantscalar");
   $result = sql($sql_key1, SCALAR, KEYED, [1]);
   push(@testres, compare(undef, $result));
   $X->{NoExec} = 0;
}

#------------------- KEYED, call errors -----------------
{
   &blurb("KEYED, no keys list");
   eval('sql("SELECT * FROM #a", HASH, KEYED)');
   push(@testres, $@ =~ /no keys/i ? 1 : 0);

   &blurb("KEYED, illegal type \$keys");
   eval('sql("SELECT * FROM #a", KEYED, undef, "a")');
   push(@testres, $@ =~ /not a .*reference/i ? 1 : 0);

   &blurb("KEYED, empty keys list");
   eval('sql("SELECT * FROM #a", HASH, KEYED, [])');
   push(@testres, $@ =~ /empty/i ? 1 : 0);

   &blurb("KEYED, undefined key name");
   eval('sql("SELECT * FROM #a", HASH, KEYED, ["bogus"])');
   push(@testres, $@ =~ /no key\b.*in result/i ? 1 : 0);

   &blurb("KEYED, key out of range");
   eval('sql("SELECT * FROM #a", LIST, KEYED, [47])');
   push(@testres, $@ =~ /number .*not valid/i ? 1 : 0);

   &blurb("KEYED, not unique");
   eval(<<'EVALEND');
       local $SIG{__WARN__} = sub {$X->cancelbatch; die $_[0]};
       sql("SELECT * FROM #a", LIST, KEYED, [1]);
EVALEND
   push(@testres, $@ =~ /not unique/i ? 1 : 0);
}

#-------------------- &callback ----------------------------
{
   my (@expect);
   my ($ix, $ok, $cancel_ix, $error_ix);
   my ($retstat);

   sub callback {
      my ($row, $ressetno) = @_;
      if ($expect[$ix][0] != $ressetno or
          not compare($row, $expect[$ix++][1])) {
         $ok = 0;
         return RETURN_CANCEL;
      }
      if ($ix == $cancel_ix) {
         return RETURN_NEXTQUERY;
      }
      if ($ix == $error_ix) {
         return RETURN_ERROR;
      }
      RETURN_NEXTROW;
   }

   &blurb("HASH, &callback");
   @expect = ([1, {a => 'A', b => 'A', i => 12}],
              [1, {a => 'A', b => 'D', i => 24}],
              [1, {a => 'A', b => 'H', i => 1}],
              [2, {sum => 37}],
              [3, {a => 'C', b => 'B', i => 12}],
              [4, {sum => 12}],
              [5, {sum => 49}],
              [6, {'x' => 'xyz'}],
              [6, {'x' => undef}],
              [8, {'Col 1' => 4711}]);
   $ix = 0;
   $cancel_ix = 0;
   $error_ix = 0;
   $ok = 1;
   $retstat = sql($sql, \&callback);
   if ($ok == 1 and $ix == $#expect + 1 and $retstat == RETURN_NEXTROW) {
      push(@testres, 1);
   }
   else {
      push(@testres, 0);
   }

   &blurb("LIST, &callback");
   @expect = ([1, ['A', 'A', 12]],
              [1, ['A', 'D', 24]],
              [2, [37]],
              [3, ['C', 'B', 12]],
              [4, [12]],
              [5, [49]],
              [6, ['xyz']],
              [6, [undef]],
              [8, [4711]]);
   $ix = 0;
   $cancel_ix = 2;
   $error_ix = 0;
   $ok = 1;
   $retstat = sql($sql, LIST, \&callback);
   if ($ok == 1 and $ix == $#expect + 1 and $retstat == RETURN_NEXTROW) {
      push(@testres, 1);
   }
   else {
      push(@testres, 0);
   }

   $ix = 0;
   $cancel_ix = 0;
   $error_ix = 3;
   $ok = 1;
   &blurb("SCALAR, &callback");
   @expect = ([1, 'A@!@A@!@12'],
              [1, 'A@!@D@!@24'],
              [1, 'A@!@H@!@1']);
   $retstat = sql($sql, \&callback, SCALAR);
   if ($ok == 1 and $ix == $#expect + 1 and $retstat == RETURN_ERROR) {
      push(@testres, 1);
   }
   else {
      push(@testres, 0);
   }

   $ix = 0;
   $cancel_ix = 0;
   $error_ix = 2;
   $ok = 1;
   blurb("sql_sp, callback");
   @expect = ([1, 'apple@!@X@!@1@!@@!@@!@T'],
              [1, 'apple@!@X@!@2@!@-15@!@@!@T']);
   $retstat = sql_sp('#sql_key_many', \&callback, SCALAR);
   if ($ok == 1 and $ix == $#expect + 1 and $retstat == RETURN_ERROR) {
      push(@testres, 1);
   }
   else {
      push(@testres, 0);
   }
}



#------------------------ Various style-parameter erros.
{
   &blurb("Bogus row style 1");
   eval('sql("SELECT * FROM #a", -23, KEYED)');
   push(@testres, $@ =~ /Illegal row.* -23 at/i ? 1 : 0);

   &blurb("Bogus row style 2");
   eval('sql("SELECT * FROM #a", undef, -23)');
   push(@testres, $@ =~ /Illegal row.* -23 at/i ? 1 : 0);

   &blurb("Bogus row style 3");
   eval('sql("SELECT * FROM #a", SINGLESET, -23)');
   push(@testres, $@ =~ /Illegal row.* -23 at/i ? 1 : 0);

   &blurb("Bogus result style");
   eval('sql("SELECT * FROM #a", LIST, -23)');
   push(@testres, $@ =~ /Illegal result.* -23 at/i ? 1 : 0);

   &blurb("Two row styles");
   eval('sql("SELECT * FROM #a", LIST, HASH)');
   push(@testres, $@ =~ /Illegal result.* 93 at/i ? 1 : 0);

   &blurb("Two result styles");
   eval('sql("SELECT * FROM #a", SINGLESET, MULTISET)');
   push(@testres, $@ =~ /Illegal row.* 139 at/i ? 1 : 0);
}

my $no_of_tests = 4 * 6 * 4 +  # Three resultstyles with result, empty, error and Noexec.
                  2 *6 +       # Two extra for SINGLEROW
                  6 + 6 +      # Extra test for KEYED (mulitple, errors, NoEexec)
                  13 + 11 +    # sql_one, seven regular + six error, Exec/NoExec
                  6 + 4 +      # 6 NORESULT  + 3 callback.
                  6;           # Style errors.
print "1..$no_of_tests\n";

my $ix = 1;
my $blurb = "";
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
