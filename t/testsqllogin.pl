#---------------------------------------------------------------------
# $Header: /Perl/OlleDB/t/testsqllogin.pl 6     08-05-01 10:48 Sommar $
#
# This file is C<required> by all test scripts. It defines a sub that
# connects to SQL Server, and changes current directory to the test
# directory, so that auxillary files are found, and all output files
# are written there.
#
# $History: testsqllogin.pl $
# 
# *****************  Version 6  *****************
# User: Sommar       Date: 08-05-01   Time: 10:48
# Updated in $/Perl/OlleDB/t
# Added parameter to permit test scripts to run without a default handle.
#
# *****************  Version 5  *****************
# User: Sommar       Date: 08-03-23   Time: 23:28
# Updated in $/Perl/OlleDB/t
# Handle empty provider value, so that it does not yield warnings about
# not being numeric.
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
   my ($use_sql_init) = @_;

   if (not defined $use_sql_init) {
      $use_sql_init = 1;
   }

   # Crack the envioronment vairable.
   my ($login) = $ENV{'OLLEDBTEST'};
   my ($server, $user, $pw, $dummy, $provider);
   ($server, $user, $pw, $dummy, $dummy, $dummy, $provider) =
        split(/;/, $login) if defined $login;
   undef $provider if defined $provider and $provider !~ /\S/;

   if ($use_sql_init) {
      return sql_init($server, $user, $pw, "tempdb", $provider);
   }
   else {
      my $X = new Win32::SqlServer;
      $X->{Provider} = $provider if defined $provider;
      $X->setloginproperty('Server', $server);
      if ($user) {
         $X->setloginproperty('Username', $user);
         $X->setloginproperty('Password', $pw);
      }
      $X->setloginproperty('Database', 'tempdb');
      $X->connect();
      return $X;
  }
}

chdir dirname($0);
if (not -d 'output') {
   mkdir('output') or die "Cannot mkdir 'output': $!\n"
}
chdir 'output';

1;
