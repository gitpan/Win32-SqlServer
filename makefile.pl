#---------------------------------------------------------------------
# $Header: /Perl/OlleDB/makefile.pl 19    08-05-12 22:04 Sommar $
#
# Makefile.pl for MSSQL::OlleDB. Note that you may need to specify where
# you ave the include files for OLE DB.
#
# $History: makefile.pl $
# 
# *****************  Version 19  *****************
# User: Sommar       Date: 08-05-12   Time: 22:04
# Updated in $/Perl/OlleDB
# Exit 0 instead of die when things are bad.
#
# *****************  Version 18  *****************
# User: Sommar       Date: 08-05-04   Time: 18:36
# Updated in $/Perl/OlleDB
# Must use delayed inmport for SQLNCLI.DLL so that Win32::SqlServer can
# run on machines without it.
#
# *****************  Version 17  *****************
# User: Sommar       Date: 08-04-28   Time: 23:10
# Updated in $/Perl/OlleDB
# Need to cater for the 64-bit version of the link library for SQL Native
# Client.
#
# *****************  Version 16  *****************
# User: Sommar       Date: 08-03-23   Time: 19:30
# Updated in $/Perl/OlleDB
# Added common debug flag. Sometimes needed when testing...
#
# *****************  Version 15  *****************
# User: Sommar       Date: 08-01-05   Time: 0:24
# Updated in $/Perl/OlleDB
# New file tableparam.obj
#
# *****************  Version 14  *****************
# User: Sommar       Date: 07-12-24   Time: 21:38
# Updated in $/Perl/OlleDB
# The big explosion: split the single source file for XS into a whole
# bunch.
#
# *****************  Version 13  *****************
# User: Sommar       Date: 07-11-26   Time: 22:46
# Updated in $/Perl/OlleDB
# Added a libfile, since we now need to link against sqlnlic10.lib.
#
# *****************  Version 12  *****************
# User: Sommar       Date: 07-09-09   Time: 0:11
# Updated in $/Perl/OlleDB
# Get SQLNCLI from 100/SDK, so that we can suppor Katmai.
#
# *****************  Version 11  *****************
# User: Sommar       Date: 07-07-10   Time: 21:14
# Updated in $/Perl/OlleDB
# New machine, new location for WINZIP.
#
# *****************  Version 10  *****************
# User: Sommar       Date: 06-04-17   Time: 21:52
# Updated in $/Perl/OlleDB
# Now forcgin the C run-time to be staically linked.
#
# *****************  Version 9  *****************
# User: Sommar       Date: 05-11-26   Time: 23:47
# Updated in $/Perl/OlleDB
# Renamed the module to Win32::SqlServer and advanced to version 2.001.
#
# *****************  Version 8  *****************
# User: Sommar       Date: 05-11-20   Time: 19:31
# Updated in $/Perl/OlleDB
# Look for SQLNCLI.H on any disk. Use -P to Winzip for correct packaging.
#
# *****************  Version 7  *****************
# User: Sommar       Date: 05-11-19   Time: 21:26
# Updated in $/Perl/OlleDB
#
# *****************  Version 6  *****************
# User: Sommar       Date: 05-11-13   Time: 16:32
# Updated in $/Perl/OlleDB
# Use /O2 for optimization.
#
# *****************  Version 5  *****************
# User: Sommar       Date: 05-07-03   Time: 23:41
# Updated in $/Perl/OlleDB
# Now we use SQLNCLI.H, which means that we will have to move away from
# VC6.
#
# *****************  Version 4  *****************
# User: Sommar       Date: 04-08-23   Time: 22:49
# Updated in $/Perl/OlleDB
#
# *****************  Version 3  *****************
# User: Sommar       Date: 04-08-23   Time: 21:52
# Updated in $/Perl/OlleDB
#
# *****************  Version 2  *****************
# User: Sommar       Date: 04-04-27   Time: 22:32
# Updated in $/Perl/MSSQL/OlleDB
#
# *****************  Version 1  *****************
# User: Sommar       Date: 04-03-18   Time: 20:24
# Created in $/Perl/MSSQL/OlleDB
#---------------------------------------------------------------------


use strict;

use ExtUtils::MakeMaker;

# Run CL to see if we are running some version of the Visual C++ compiler.
my $cl = `cl 2>&1`;
my $clversion = 0;
if ($cl =~ m!^Microsoft.*C/C\+\+\s+Optimizing\s+Compiler\s+Version\s+(\d+)!i) {
   $clversion = $1;
}

if ($clversion == 0) {
   warn "You don't appear to have Visual C++ installed. If you use another\n";
   warn "C++ compiler, I have no idea whether that will work or not. Be warned!\n";
}
elsif ($clversion < 13) {
   warn "You are using Visual C++ 6.0 or earlier. Unfortunately, OlleDB.xs\n";
   warn "performs an #include of SQLNCLI.H which does not compile with VC6.\n";
   warn  "No MAKEFILE generated.\n";
   exit 0
}

my $SQLDIR  = '\Program Files\Microsoft SQL Server\100\SDK';
my $sqlnclih = "$SQLDIR\\INCLUDE\\SQLNCLI.H";
foreach my $device ('C'..'Z') {
   if (-r "$device:$sqlnclih") {
      $SQLDIR = "$device:$SQLDIR";
      last;
   }
}
if ($SQLDIR !~ /^[C-Z]:/) {
    warn "Can't find '$sqlnclih' on any disk.\n";
    warn 'Check setting of $SQLDIR in makefile.pl' . "\n";
    warn "No MAKEFILE generated.\n";
    exit 0;
}

my $archlibdir = ($ENV{PROCESSOR_ARCHITECTURE} eq 'AMD64' ? 'x64' : $ENV{PROCESSOR_ARCHITECTURE});
my $libfile = qq!"$SQLDIR\\LIB\\$archlibdir\\sqlncli10.lib"!;

WriteMakefile(
    'INC'          => ($SQLDIR ? qq!-I"$SQLDIR\\INCLUDE"! : ""),
    'NAME'         => 'Win32::SqlServer',
    'CCFLAGS'     => '/MT', # Debug-flag: '/Zi',
    'OPTIMIZE'     => '/O2',
    'OBJECT'       => 'SqlServer.obj handleattributes.obj convenience.obj ' .
                      'datatypemap.obj init.obj internaldata.obj ' .
                      'errcheck.obj connect.obj utils.obj datetime.obj ' .
                      'tableparam.obj senddata.obj getdata.obj',
    'LIBS'         => [":nosearch :nodefault $libfile delayimp.lib kernel32.lib user32.lib ole32.lib oleaut32.lib uuid.lib libcmt.lib"],
    'VERSION_FROM' => 'SqlServer.pm',
    'XS'           => { 'SqlServer.xs' => 'SqlServer.cpp' },
    'dist'         => {ZIP => '"C:\Program Files (x86)\Winzip\wzzip"',
                       ZIPFLAGS => '-r -P'},
    'dynamic_lib'  => { OTHERLDFLAGS => '/base:"0x19860000" /DELAYLOAD:sqlncli10.dll'}
    # Set base address to avoid DLL collision, makes startup speedier. Remove
    # if your compiler don't have this option.
);

sub MY::xs_c {
    '
.xs.cpp:
   $(PERL) -I$(PERL_ARCHLIB) -I$(PERL_LIB) $(XSUBPP) $(XSPROTOARG) $(XSUBPPARGS) $*.xs >xstmp.c && $(MV) xstmp.c $*.cpp
';
}

