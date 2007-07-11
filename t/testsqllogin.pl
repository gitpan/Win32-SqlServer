#---------------------------------------------------------------------
# $Header: /Perl/OlleDB/t/testsqllogin.pl 4     07-07-07 16:43 Sommar $
#
# This file is C<required> by all test scripts. It defines a sub that
# connects to SQL Server, and changes current directory to the test
# directory, so that auxillary files are found, and all output files
# are written there.
#
# $History: testsqllogin.pl $
# 
# *****************  Version 4  *****************
# User: Sommar       Date: 07-07-07   Time: 16:43
# Updated in $/Perl/OlleDB/t
# Added support for specifying differnt providers.
#
# *****************  Version 3  *****************
# User: Sommar       Date: 05-07-16   Time: 17:30
# Updated in $/Perl/OlleDB/t
# We now have all the action in a special output directory.
#
# *****************  Version 2  *****************
# User: Sommar       Date: 05-06-27   Time: 21:40
# Updated in $/Perl/OlleDB/t
# Change directory to the test directory.
#
# *****************  Version 1  *****************
# User: Sommar       Date: 05-01-02   Time: 20:54
# Created in $/Perl/OlleDB/t
#
#---------------------------------------------------------------------


sub testsqllogin
{
   my ($login) = $ENV{'OLLEDBTEST'};
   my ($server, $user, $pw, $dummy, $provider);
   ($server, $user, $pw, $dummy, $dummy, $dummy, $provider) =
        split(/;/, $login) if defined $login;
   return sql_init($server, $user, $pw, "tempdb", $provider);
}

chdir dirname($0);
if (not -d 'output') {
   mkdir('output') or die "Cannot mkdir 'output': $!\n"
}
chdir 'output';

1;
