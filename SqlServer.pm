#---------------------------------------------------------------------
# $Header: /Perl/OlleDB/SqlServer.pm 42    06-04-17 21:48 Sommar $
#
# Copyright (c) 2004-2006 Erland Sommarskog
#
#
# $History: SqlServer.pm $
# 
# *****************  Version 42  *****************
# User: Sommar       Date: 06-04-17   Time: 21:48
# Updated in $/Perl/OlleDB
# Advancrd version to 2.002. No other changes.
#
# *****************  Version 41  *****************
# User: Sommar       Date: 05-11-26   Time: 23:47
# Updated in $/Perl/OlleDB
# Renamed the module to Win32::SqlServer and advanced to version 2.001.
#
# *****************  Version 40  *****************
# User: Sommar       Date: 05-11-13   Time: 16:33
# Updated in $/Perl/OlleDB
#
#---------------------------------------------------------------------


package Win32::SqlServer;

require 5.008003;

use strict;
use Exporter;
use DynaLoader;
use Tie::Hash;
use Carp;

use vars qw(@ISA @EXPORT @EXPORT_OK %EXPORT_TAGS
            $def_handle $SQLSEP
            %VARLENTYPES %STRINGTYPES %QUOTEDTYPES %UNICODETYPES %LARGETYPES
            %BINARYTYPES %DECIMALTYPES %MAXTYPES %TYPEINFOTYPES $VERSION);


$VERSION = '2.002';

@ISA = qw(Exporter DynaLoader Tie::StdHash);

# Kick life into the C++ code.
bootstrap Win32::SqlServer;

@EXPORT = qw(sql_init sql_string);
@EXPORT_OK = qw(sql_set_conversion sql_unset_conversion sql_one sql sql_sp
                sql_insert sql_has_errors sql_get_command_text
                sql_begin_trans sql_commit sql_rollback
                NORESULT SINGLEROW SINGLESET MULTISET KEYED
                SCALAR LIST HASH
                $SQLSEP
                TO_SERVER_ONLY TO_CLIENT_ONLY TO_SERVER_CLIENT
                RETURN_NEXTROW RETURN_NEXTQUERY RETURN_CANCEL RETURN_ERROR
                RETURN_ABORT
                PROVIDER_DEFAULT PROVIDER_SQLNCLI PROVIDER_SQLOLEDB
                DATETIME_HASH DATETIME_ISO DATETIME_REGIONAL DATETIME_FLOAT
                DATETIME_STRFMT
                CMDSTATE_INIT CMDSTATE_ENTEREXEC CMDSTATE_NEXTRES
                CMDSTATE_NEXTROW CMDSTATE_GETPARAMS);
%EXPORT_TAGS = (consts       => [qw($SQLSEP)],   # Filled in below.
                routines     => [qw(sql_set_conversion sql_unset_conversion
                                    sql_one sql sql_sp sql_insert
                                    sql_has_errors sql_get_command_text
                                    sql_string
                                    sql_begin_trans sql_commit sql_rollback)],
                resultstyles => [qw(NORESULT SINGLEROW SINGLESET MULTISET KEYED)],
                rowstyles    => [qw(SCALAR LIST HASH)],
                directions   => [qw(TO_SERVER_ONLY TO_CLIENT_ONLY TO_SERVER_CLIENT)],
                returns      => [qw(RETURN_NEXTROW RETURN_NEXTQUERY RETURN_CANCEL
                                    RETURN_ERROR RETURN_ABORT)],
                providers    => [qw(PROVIDER_DEFAULT PROVIDER_SQLNCLI PROVIDER_SQLOLEDB)],
                datetime     => [qw(DATETIME_HASH DATETIME_ISO DATETIME_REGIONAL
                                    DATETIME_FLOAT DATETIME_STRFMT)],
                cmdstates    => [qw(CMDSTATE_INIT CMDSTATE_ENTEREXEC CMDSTATE_NEXTRES
                                    CMDSTATE_NEXTROW CMDSTATE_GETPARAMS)]
                );
push(@{$EXPORT_TAGS{'consts'}}, @{$EXPORT_TAGS{'routines'}},
                                @{$EXPORT_TAGS{'resultstyles'}},
                                @{$EXPORT_TAGS{'rowstyles'}},
                                @{$EXPORT_TAGS{'directions'}},
                                @{$EXPORT_TAGS{'returns'}},
                                @{$EXPORT_TAGS{'providers'}},
                                @{$EXPORT_TAGS{'datetime'}});

# Result-style constants.
use constant NORESULT  => 821;
use constant SINGLEROW => 741;
use constant SINGLESET => 643;
use constant MULTISET  => 139;
use constant KEYED     => 124;
use constant RESULTSTYLES => (NORESULT, SINGLEROW, SINGLESET, MULTISET, KEYED);

# Row-style constants.
use constant SCALAR    => 17;
use constant LIST      => 89;
use constant HASH      => 93;
use constant ROWSTYLES => (SCALAR, LIST, HASH);

# Separator when rows returned in one string, reconfigurarable.
$SQLSEP = "\022";

# Constants for conversion direction
use constant TO_SERVER_ONLY    => 8798;
use constant TO_CLIENT_ONLY    => 3456;
use constant TO_SERVER_CLIENT  => 2402;

# Constants for return values for callbacks
use constant RETURN_NEXTROW    =>  1;
use constant RETURN_NEXTQUERY  =>  2;
use constant RETURN_CANCEL     =>  3;
use constant RETURN_ERROR      =>  0;
use constant RETURN_ABORT      => -1;

# Constants for option Provider
use constant PROVIDER_DEFAULT  => 0;
use constant PROVIDER_SQLOLEDB => 1;
use constant PROVIDER_SQLNCLI  => 2;
use constant PROVIDER_OPTIONS => (PROVIDER_DEFAULT, PROVIDER_SQLOLEDB,
                                  PROVIDER_SQLNCLI);

# Constants for datetime options
use constant DATETIME_HASH     => 0;
use constant DATETIME_ISO      => 1;
use constant DATETIME_REGIONAL => 2;
use constant DATETIME_FLOAT    => 3;
use constant DATETIME_STRFMT   => 4;
use constant DATETIME_OPTIONS  => (DATETIME_HASH, DATETIME_ISO,
                                   DATETIME_REGIONAL, DATETIME_FLOAT,
                                   DATETIME_STRFMT);

# Constants for command state.
use constant CMDSTATE_INIT      => 0;
use constant CMDSTATE_ENTEREXEC => 1;
use constant CMDSTATE_NEXTRES   => 2;
use constant CMDSTATE_NEXTROW   => 3;
use constant CMDSTATE_GETPARAMS => 4;

use constant PACKAGENAME => 'Win32::SqlServer';

# Constant hashes for datatype combinations, for internal use only.
%VARLENTYPES  = ('char' => 1, 'nchar' => 1, 'varchar' => 1, 'nvarchar' => 1,
                 'binary' => 1, 'varbinary' => 1, 'UDT' => 1);
%STRINGTYPES  = ('char' => 1, 'varchar' => 1, 'nchar' => 1, 'nvarchar' => 1,
                 'xml' => 1, 'text'=> 1, 'ntext' => 1);
%LARGETYPES   = ('text' => 1, 'ntext' => 1, 'image' => 1, 'xml' => 1);
%QUOTEDTYPES  = ('char' => 1, 'varchar' => 1, 'nchar' => 1, 'nvarchar' => 1,
                 'text' => 1, 'ntext' => 1, 'uniqueidentifier' => 1,
                 'datetime' => 1 , 'smalldatetime'=> 1);
%UNICODETYPES  = ('nchar' => 1, 'nvarchar' => 1, 'ntext' => 1);
%BINARYTYPES   = ('binary' => 1, 'varbinary' => 1, 'timestamp' => 1,
                  'image' => 1, 'UDT' => 1);
%DECIMALTYPES  = ('decimal' => 1, 'numeric' => 1);
%MAXTYPES      = ('varchar' => 1, 'nvarchar' => 1, 'varbinary' =>1);
%TYPEINFOTYPES = ('UDT' => 1, 'xml' => 1);

# Global hash to keep track of all object we create and destroy. This is
# needed when cloning for a new thread.
my %my_objects;

#----- -------------- Set up supported attributes. --------------------------
my %myattrs;

use constant XS_ATTRIBUTES =>   # Used by the XS code.
             qw(internaldata Provider PropsDebug AutoConnect RowsAtATime
                DecimalAsStr DatetimeOption BinaryAsStr DateFormat MsecFormat
                CommandTimeout MsgHandler QueryNotification);
use constant PERL_ATTRIBUTES => # Attributes used by the Perl code.
             qw(ErrInfo SQL_version to_server to_client NoExec procs tables
                LogHandle UserData);
use constant ALL_ATTRIBUTES => (XS_ATTRIBUTES, PERL_ATTRIBUTES);

foreach my $attr (ALL_ATTRIBUTES) {
   $myattrs{$attr}++;
}

#------------------------  FETCH and STORE -------------------------------
# My own FETCH routine, chckes that retrieval is of a known attribute.
sub FETCH {
   my ($self, $key) = @_;
   if (not exists $myattrs{$key}) {
       # Compability with MSSQL::Sqllib: permit initial lowercase.
       $key =~ s/^./uc($&)/e;
       if (not exists $myattrs{$key}) {
           $self->olle_croak("Attempt to fetch a non-existing Win32::SqlServer property '$key'");
       }
   }
   if ($key eq "SQL_version" and not defined $self->{$key}) {
       # If don't have it, we must retrieve it, and save it. There is a
       # special routine for this.
       $self->{SQL_version} = $self->get_sqlserver_version;
   }

   unless ($key eq 'Provider') {
      return $self->{$key};
   }
   else {
      return $self->get_provider_enum;
   }
}

# My own STORE routine, barfs if attribute is non-existent.
sub STORE {
   my ($self, $key, $value) = @_;
   if (not exists $myattrs{$key}) {
       $key =~ s/^./uc($&)/e;
       if (not exists $myattrs{$key}) {
           $self->olle_croak("Attempt to set a non-existing Win32::SqlServer property '$key'");
       }
   }
   my $old_value = $self->{$key};
   if ($key eq 'MsgHandler') {
      if ($value) {
         if (not ref $value eq "CODE") {
            # It is not a ref to a sub, but it could be the name of that. There
            # is an XS routine to validate this. It croaks if things are bad.
            $self->validatecallback($value);
         }
      }
      else {
         $value = undef;
      }
   }
   elsif ($key eq "internaldata" or $key eq "ErrInfo") {
      if ($old_value) {
         my $caller = (caller(1))[3];
         unless ($caller and $caller eq PACKAGENAME . '::DESTROY') {
            $self->olle_croak("You must not change the object property '$key'");
         }
      }
   }
   elsif ($key eq "Provider") {
      if (not grep($value == $_, PROVIDER_OPTIONS)) {
         $self->olle_croak("Illegal value '$value' for the Provider property");
      }
      my $ret = $self->set_provider_enum($value);
      if ($ret == -1) {
         croak("Cannot set the Provider while connected");
      }
   }
   elsif ($key eq "DatetimeOption") {
      if (not grep($value == $_, DATETIME_OPTIONS)) {
         $self->olle_croak("Illegal value '$value' for the DatetimeOption property");
      }
   }
   elsif ($key eq "QueryNotification") {
      if (not ref $value eq "HASH") {
         $self->olle_croak("The value for the QueryNotification property must be a hash reference");
      }
   }

   $self->{$key} = $value;
}

sub DELETE {
   # Generally it is not permitted to delete keys from the hash, but there
   # is an exception for SQL_version, since the XS version needs to clear it,
   # but for some reason is not permitted to write to the hash... Also,
   # to_server and to_client are deleted by sql_unset_conversion.
   my ($self, $key) = @_;
   if (not grep($_ eq $key, qw(SQL_version to_server to_client))) {
      $self->olle_croak ("Attempt to delete the object property '$key'");
   }
   $self->{$key} = undef;
}

#------------------------ New and DESTROY -------------------------------

sub new {
   my ($self) = @_;

   my (%olle);

   # %olle is our tied hash.
   my $X = tie %olle, $self;

   # Save a reference that we created it. This is for CLONE, see below.
   $my_objects{$X} = $X;

   # Initiate Win32::SqlServer properties.
   $olle{"internaldata"}      =  setupinternaldata();
   $olle{"AutoConnect"}       = 0;
   $olle{"PropsDebug"}        = 0;
   $olle{"RowsAtATime"}       = 100;
   $olle{"DecimalAsStr"}      = 0;
   $olle{"DatetimeOption"}    = DATETIME_ISO;
   $olle{"BinaryAsStr"}       = '1';
   $olle{"DateFormat"}        = "%Y%m%d %H:%M:%S";
   $olle{"MsecFormat"}        =  ".%3.3d";
   $olle{"CommandTimeout"}    = 0;
   $olle{"QueryNotification"} = {};
   $olle{"MsgHandler"}        = \&sql_message_handler;

   # Initiate error handling.
   $olle{ErrInfo} = new_err_info();

   # Return the blessed object.
   bless \%olle, PACKAGENAME;
}

sub CLONE {
# Perl calls this routine when a new thread is created. If we would do
# nothing at all, internaldata would be the same for all thread, which
# would only cause misery. Particularly, the child threads would try to
# deallocate it, which crashes with "attempt to free from wrong pool".
# So we give all cloned objects a new fresh internaldata.
   foreach my $obj (values %my_objects) {
      $$obj{"internaldata"} = setupinternaldata()
   }
}

sub DESTROY {
   my ($self) = @_;

   delete $my_objects{$self};

   # We run the destruction in eval, as Perl sometimes produces an error
   # message "Can't call method "FETCH" on an undefined value" when the
   # destructor is called a second time.
   eval('xs_DESTROY($self)');

   unless ($@) {
      # We must clear internaldata, since Perl calls the destructor twice, but
      # on the second occasion, the XS code has already deallocated internaldata.
      # The XS code has problem with setting values in stored hashes, why we do
      # it. This assignment cannot be in eval, since the STORE method only
      # permits DESTROY to change internaldata.
      $$self{'internaldata'} = 0;
   }
}

#--------------------  sql_init  ----------------------------------------
sub sql_init {
# Logs into SQL Server and returns an object to use for further communication
# with the module.
    my ($server, $user, $pw, $db) = @_;

    my $X = new(PACKAGENAME);

    # Set login properties if provided.
    $X->setloginproperty('Server', $server) if $server;

    if ($user) {
       $X->setloginproperty('Username', $user);
       $X->setloginproperty('Password', $pw) if $pw;
    }
    $X->setloginproperty('Database', $db) if $db;

    # Login into the server.
    if (not $X->connect()) {
       croak("Login into SQL Server failed");
    }

    # Get SQL version.
    $X->{SQL_version} = $X->get_sqlserver_version();

    # If the global default handle is undefined, give the recently created
    # connection.
    if (not defined $def_handle) {
       $def_handle = $X;
    }

    $X;
}


#-------------------------- sql_set_conversion --------------------------
sub sql_set_conversion
{
    my($X) = (ref @_[$[] eq PACKAGENAME ? shift @_ : $def_handle);
    my($client_cs, $server_cs, $direction) = @_;

    # First validate the $direction parameter.
    if (! $direction) {
       $direction = TO_SERVER_CLIENT;
    }
    if (! grep($direction == $_,
              (TO_SERVER_ONLY, TO_CLIENT_ONLY, TO_SERVER_CLIENT))) {
       $X->olle_croak("Illegal direction value: $direction");
    }

    # Normalize parameters and get defaults. The client charset.
    if (not $client_cs or $client_cs =~ /^OEM/i) {
       # No value or OEM, read actual OEM codepage from registry.
       $client_cs = get_codepage_from_reg('OEMCP');
    }
    elsif ($client_cs =~ /^ANSI$/i) {
       # Read ANSI code page.
       $client_cs = get_codepage_from_reg('ACP');
    }
    $client_cs =~ s/^cp_?//i;             # Strip CP[_]
    if ($client_cs =~ /^\d{3,3}$/) {
       $client_cs = "0$client_cs";       # Add leading zero.
    }

    # Now the server charset. If no charset given, query the server.
    if (not $server_cs) {
       if ($X->{SQL_version} =~ /^[467]\./) {
          # SQL Server 7.0 or earlier.
          $server_cs = $X->internal_sql(<<SQLEND, SCALAR, SINGLEROW);
               SELECT chs.name
               FROM   master..syscharsets sor, master..syscharsets chs,
                      master..syscurconfigs cfg
               WHERE  cfg.config = 1123
                 AND  sor.id     = cfg.value
                 AND  chs.id     = sor.csid
SQLEND
       }
       else {
          # Modern stuff, SQL 2000 or later.
          $server_cs = $X->internal_sql(<<SQLEND, SCALAR, SINGLEROW);
             SELECT collationproperty(
                    CAST(serverproperty ('collation') as nvarchar(255)),
                    'CodePage')
SQLEND
       }
    }
    if ($server_cs =~ /^iso_1$/i) {    # iso_1 is how SQL6&7 reports Latin-1.
       $server_cs = "1252";            # CP1252 is the Latin-1 code page.
    }
    $server_cs =~ s/^cp_?//i;
    if ($server_cs =~ /^\d{3,3}$/) {
       $server_cs = "0$server_cs";
    }

    # If client and server charset are the same, we should only remove any
    # current conversion, and then quit.
    if ("\U$client_cs\E" eq "\U$server_cs\E") {
       $X->sql_unset_conversion($direction);
       return;
    }

    # Now we try to find a file in System32.
    my($server_first) = 1;
    my($server_first_name) = "$ENV{'SYSTEMROOT'}\\System32\\$server_cs$client_cs.cpx";
    my($client_first_name) = "$ENV{'SYSTEMROOT'}\\System32\\$client_cs$server_cs.cpx";
    if (not open(F, $server_first_name)) {
       open(F, $client_first_name) or
          $X->olle_croak("Can't open neither '$server_first_name' nor '$client_first_name'");
       $server_first = 0;
    }

    # First read translations from the first charset. But the chars into
    # a string. When used the strings will be fed to tr.
    my($server_repl, $server_with) = ("", "");
    my($line);
    #while (($line = $F->getline) !~ m!^/!) {
    while (($line = <F>) !~ m!^/!) {
       chop $line;
       next if $line !~ /:/;
       my($a, $b) = split(/:/, $line);
       $server_repl .= chr($a);
       $server_with .= chr($b);
    }

    # The other half.
    my($client_repl, $client_with) = ("", "");
    #while ($line = $F->getline) {
    while (defined ($line = <F>)) {
       chop $line;
       next if $line !~ /:/;
       my($a, $b) = split(/:/, $line);
       $client_repl .= chr($a);
       $client_with .= chr($b);
    }

    close F;

    # Swap the strings if client's charset was first in the file.
    if (! $server_first) {
       ($client_repl, $server_repl) = ($server_repl, $client_repl);
       ($client_with, $server_with) = ($server_with, $client_with);
    }

    # Store the charset converstions into the handle. We store these as
    # subroutines ready to use. We need to use eval, as tr is static.
    if ($direction == TO_SERVER_ONLY or $direction == TO_SERVER_CLIENT) {
       $X->{'to_server'} = eval("sub { foreach (\@_) {next if ref;
                                  tr/\Q$client_repl\E/\Q$client_with\E/ if \$_}
                                 return \@_}") or
           $X->olle_croak("eval of client-to-server conversion failed: $@");
    }
    if ($direction == TO_CLIENT_ONLY or $direction == TO_SERVER_CLIENT) {
    # For server-to-client we need a return value for hashes.
       $X->{'to_client'} = eval("sub { foreach (\@_) { next if ref;
                                  tr/\Q$server_repl\E/\Q$server_with\E/ if \$_}
                                 return \@_}") or
           $X->olle_croak("eval of server-to-client conversion failed: $@");
    }
}

#-------------------------- sql_unset_conversion -------------------------
sub sql_unset_conversion
{
    my($X) = (ref @_[$[] eq PACKAGENAME ? shift @_ : $def_handle);
    my($direction) = @_;

    # First validate the $direction parameter.
    if (! $direction) {
       $direction = TO_SERVER_CLIENT;
    }
    if (! grep($direction == $_,
              (TO_SERVER_ONLY, TO_CLIENT_ONLY, TO_SERVER_CLIENT))) {
       $X->olle_croak("Illegal direction value: $direction");
    }

    # Now remove as ordered.
    if ($direction == TO_SERVER_ONLY or $direction == TO_SERVER_CLIENT) {
       delete $X->{'to_server'};
    }
    if ($direction == TO_CLIENT_ONLY or $direction == TO_SERVER_CLIENT) {
       delete $X->{'to_client'};
    }
}

#-----------------------------  sql_one-------------------------------------
sub sql_one
{
    my($X) = (ref @_[$[] eq PACKAGENAME ? shift @_ : $def_handle);
    my($sql) = shift @_;

    # Get parameter array if any.
    my ($hashparams, $arrayparams);
    if (ref $_[0] eq "ARRAY") {
       $arrayparams = shift @_;
    }
    if (ref $_[0] eq "HASH") {
       $hashparams = shift @_;
    }

    # Get rowstyle.
    my ($rowstyle) = @_;

    my ($dataref, $saveref, $exec_ok);

    # Make sure $rowstyle has a legal value.
    $rowstyle = $rowstyle || (wantarray ? HASH : SCALAR);
    check_style_params($rowstyle);

    # Apply conversion.
    $X->do_conversion('to_server', $sql);

    # Set up the command - run initbatch and enter parameters if necessary.
    my $ret = $X->setup_sqlcmd($sql, $arrayparams, $hashparams);
    if (not $ret) {
        $X->olle_croak("Single-row query '$sql' had parameter errors");
    }

    # Do logging.
    $X->do_logging;

    if ($X->{'NoExec'}) {
       $X->cancelbatch;
       return (wantarray ? () : undef);
    }

    # Run the command.
    $exec_ok = $X->executebatch;

    # Get the only result set and the only row - or at least there should
    # be exactly one of each.
    my $sets = 0;
    my $rows = 0;
    if ($exec_ok) {
       # Only try this if query executed.
       while ($X->nextresultset()) {
          $sets++;

          while ($X->nextrow(($rowstyle == HASH) ? $dataref : undef,
                             ($rowstyle == HASH) ? undef : $dataref)) {
             $rows++;
             # If we have a second row, something is wrong.
             if ($rows > 1) {
                $X->olle_croak("Single-row query '$sql' returned more than one row");
             }
             $saveref = $dataref;
          }
       }
    }

    # Buf if execution failed, we are seeing the now.
    # If we don't have any result set, something is wrong.
    $X->olle_croak("Single-row query '$sql' returned no result set") if $sets == 0;

    # Same if we have no row at at all.
    $X->olle_croak("Single-row query '$sql' returned no row") if $rows == 0;

    # Apply server-to-client conversion
    $X->do_conversion('to_client', $saveref);

    if (wantarray) {
       return (($rowstyle == HASH) ? %$saveref : @$saveref);
    }
    else {
       return (($rowstyle == SCALAR) ? list_to_scalar($saveref) : $saveref);
    }
}

#-----------------------  sql  --------------------------------------
sub sql
{
    my($X) = (ref @_[$[] eq PACKAGENAME ? shift @_ : $def_handle);

    my $sql = shift @_;

    # Get parameter array if any.
    my ($arrayparams, $hashparams);
    if (ref $_[0] eq "ARRAY") {
       $arrayparams = shift @_;
    }
    if (ref $_[0] eq "HASH") {
       $hashparams = shift @_;
    }

    # Style parameters. Get them from @_ and then check that values are
    # legal and supply defaults as needed.
    my($rowstyle, $resultstyle, $keys) = @_;
    check_style_params($rowstyle, $resultstyle, $keys);

    # Apply conversion.
    $X->do_conversion('to_server', $sql);

    # Set up the SQL command - initbatch and enter parameters if necesary.
    my $ret = $X->setup_sqlcmd($sql, $arrayparams, $hashparams);
    if (not $ret) {
       return (wantarray ? () : undef);
    }

    # Log the statement.
    $X->do_logging;

    my $exec_ok;
    unless ($X->{'NoExec'}) {
       # Run the command.
       $exec_ok = $X->executebatch;
    }
    else {
       $X->cancelbatch;
       $exec_ok = 0;
    }

    # And get the resultsets.
    return $X->do_result_sets($exec_ok, $rowstyle, $resultstyle, $keys);
}

#-------------------------- sql_sp ------------------------------------
sub sql_sp {
    my($X) = (ref @_[$[] eq PACKAGENAME ? shift @_ : $def_handle);

    # In this one we're not taking all parameters at once, but one by one,
    # as the parameter list is quite variable.
    my ($SP, $retvalueref, $unnamed, $named, $rowstyle,
        $resultstyle, $keys, $dummy);

    # The name of the SP, mandatory.
    $SP = shift @_;

    # Reference to scalar to receive the return value. Since there always is
    # return value, we always has a reference to a place to store it.
    if (ref $_[0] eq "SCALAR") {
       $retvalueref = shift @_;
    }
    else {
       $retvalueref = \$dummy;
    }

    # Reference to a array with named parameters.
    if (ref $_[0] eq "ARRAY") {
       $unnamed = shift @_;
    }

    # Reference to a hash with named parameters.
    if (ref $_[0] eq "HASH") {
       $named = shift @_;
    }

    # The usual row- and result-style parameters.
    ($rowstyle, $resultstyle, $keys) = @_;
    check_style_params($rowstyle, $resultstyle, $keys);

    # Reference to hash that holds the parameter definitions.
    my ($paramdefs);

    # If we have the parameter profile for this SP, we can reuse it.
    if (exists $X->{procs}{$SP}) {
       $paramdefs = $X->{'procs'}{$SP}{'params'};
    }
    else {
       # No we don't. We must retrieve from the server.

       # Get the object id for the table and it's database
       my ($objid, $objdb, $normalspec) = $X->get_object_id($SP);
       if (not defined $objid) {
          my $msg = "Stored procedure '$SP' is not accessible";
          $X->olledb_message(-1, 1, 16, $msg);
          return (wantarray ? () : undef);
       }

       # Now, inquire about all the parameters their types. Always include
       # the return value. It's in the system metadata only for UDFs, so for
       # SPs, we have to roll our own. Different handling for different SQL
       # Server versions.
       # The second UNION bit is for the return value from SP:s.
       my $getcols;
       if ($X->{SQL_version} =~ /^6\./) {
          $getcols = <<SQLEND;
              SELECT c.name, paramno = c.colid,
                     type = CASE c.usertype
                               WHEN 80 THEN ut.name
                               ELSE t.name
                            END,
                     max_length = c.length, "precision" = coalesce(c.prec, 0),
                     scale = coalesce(c.scale, 0), is_input = 1,
                     is_output = CASE c.status & 0x40
                                    WHEN 0 THEN 0
                                    ELSE 1
                                 END, is_retstatus = 0, typeinfo = NULL
              FROM   $objdb.dbo.syscolumns c
              JOIN   $objdb.dbo.systypes ut ON c.usertype = ut.usertype
              JOIN   $objdb.dbo.systypes t  ON ut.type = t.type
              WHERE  c.id = ?
                AND  t.usertype < 80
                AND  t.name <> 'sysname'
             UNION
             SELECT  NULL, 0, 'int', 4, 0, 0, 0, 1, 1, NULL
             ORDER   BY paramno
SQLEND
       }
       elsif ($X->{SQL_version} =~ /^[78]\./) {
          # The CASE for is_output because SQL 2000 says 0 for ret value from UDF.
          $getcols = <<SQLEND;
              SELECT name = CASE colid WHEN 0 THEN NULL ELSE name END,
                     paramno = colid, type = type_name(xtype),
                     max_length = length, "precision" = coalesce(prec, 0),
                     scale = coalesce(scale, 0),
                     is_input  = CASE colid WHEN 0 THEN 0 ELSE 1 END,
                     is_output = CASE colid WHEN 0 THEN 1 ELSE isoutparam END,
                     is_retstatus = 0, typeinfo = NULL
              FROM   $objdb.dbo.syscolumns
              WHERE  id = \@objid
              UNION
              SELECT NULL, 0, 'int', 4, 0, 0, 0, 1, 1, NULL
              WHERE  NOT EXISTS (SELECT *
                                 FROM   $objdb.dbo.syscolumns
                                 WHERE  id = \@objid
                                   AND  colid = 0)
              ORDER   BY paramno
SQLEND
       }
       else {
          # SQL Server 2005 or later.
          $getcols = <<SQLEND;
              SELECT name = CASE p.parameter_id WHEN 0 THEN NULL ELSE p.name END,
                     paramno = p.parameter_id,
                     type = CASE p.system_type_id
                               WHEN 240 THEN 'UDT'
                               ELSE type_name(p.system_type_id)
                          END,
                     p.max_length, p.precision, p.scale,
                     is_input = CASE p.parameter_id WHEN 0 THEN 0 ELSE 1 END,
                     p.is_output, is_retstatus = 0,
                     typeinfo =
                     CASE p.system_type_id
                          WHEN 240
                          THEN \@objdb + '.' + quotename(s1.name) + '.' +
                                quotename(t.name)
                          WHEN 241
                          THEN \@objdb + '.' + quotename(s2.name) + '.' +
                               quotename(x.name)
                     END
              FROM   $objdb.sys.all_parameters p
              LEFT   JOIN ($objdb.sys.types t
                          JOIN  $objdb.sys.schemas s1 ON t.schema_id = s1.schema_id)
                  ON  p.user_type_id = t.user_type_id
                 AND  t.is_assembly_type = 1
              LEFT   JOIN ($objdb.sys.xml_schema_collections x
                           JOIN  $objdb.sys.schemas s2 ON x.schema_id = s2.schema_id)
                  ON  p.xml_collection_id = x.xml_collection_id
              WHERE  object_id = \@objid
              UNION
              SELECT NULL, 0, 'int', 4, 0, 0, 0, 1, 1, NULL
              WHERE  NOT EXISTS (SELECT *
                                 FROM   $objdb.sys.all_parameters
                                 WHERE  object_id = \@objid
                                   AND  parameter_id = 0)
              ORDER   BY paramno
SQLEND
       }

       # Trim the SQL from extraneous spaces, to save network bandwidth.
       $getcols =~ s/\s{2,}/ /g;

       # Get the data. 6.5 has a special call since it does not support
       # named parameters.
       if ($X->{SQL_version} =~ /^6\./) {
          $paramdefs = $X->internal_sql($getcols, [['int', $objid]], HASH);
       }
       else {
          $paramdefs = $X->internal_sql($getcols,
                                       {'@objid' => ['int',      $objid],
                                        '@objdb' => ['nvarchar', $objdb]},
                                        HASH);
       }

       # Remove irrelevant statement text.
       undef $X->{ErrInfo}{SP_call};

       # Store the profile in the handle.
       $X->{'procs'}{$SP}{'params'} = $paramdefs;
       $X->{'procs'}{$SP}{'normal'} = $normalspec;
    }

    # Check that the number of unnamed parameters does not exceed the
    # number of parameters the SP actually have.
    if ($unnamed and $#$unnamed > $#$paramdefs - 1) {
       my $no_of_passed = $#$unnamed + 1;
       my $no_of_real = $#$paramdefs;   # Since @paramdefs include return value.
       my $msg = ($no_of_passed > 1 ?
                 "There were $no_of_passed parameters " :
                 "There was a parameter ") .
                 "passed for procedure '$SP' that does ";
       if ($no_of_real == 0) {
          $msg .= "not take any parameters.";
       }
       elsif ($no_of_real == 1) {
          $msg .= "only take one parameter.";
       }
       else {
          $msg .= "only take $no_of_real parameters.";
       }
       $X->olledb_message(-1, 1, 16, $msg);
       return (wantarray ? () : undef);
    }

    # At this point we need one array for parameters, and one to receive
    # parameters.
    my($no_of_pars, @all_parameters, @output_params);

    # The return value is first in line.
    $no_of_pars = 0;
    push(@all_parameters, \$retvalueref);

    # Copy a reference for all unnamed parameters.
    foreach my $ix (0..$#$unnamed) {
       push(@all_parameters, \$$unnamed[$ix]);
    }
    $no_of_pars += scalar(@$unnamed);

    # And put named parameters on the slot they appear in the parameter
    # list.
    if ($named and %$named) {
       # Get a crossref from name to position.
       my (%crossref, $no_of_errs);
       foreach my $param (@$paramdefs) {
          $crossref{$$param{'name'}} = $$param{'paramno'}
              if defined $$param{'name'};
       }

       foreach my $key (keys %$named) {
          my $name = $key;

          # Add '@' if missing, but check for duplicates.
          if ($name !~ /^\@/) {
             if (exists($$named{'@' . $key})) {
                my $msg = "Warning: hash parameters for '$SP' includes the key " .
                          "'$name' as well as '\@$name'. The value for '$name' " .
                          "is discarded.";
                $X->olledb_message(-1, 1, 10, $msg);
                next;
             }
             $name = '@' . $name;
          }

          # Check that there is such a parameter
          if (not exists $crossref{$name}) {
             my $msg = "Procedure '$SP' does not have a parameter '$name'";
             $X->olledb_message(-1, 1, 10, $msg);
             $no_of_errs++;
             next;
          }

          my $parno = $crossref{$name};

          if (defined $all_parameters[$parno] and $^W) {
             my $msg = "Parameter '$name' in position $parno for '$SP' " .
                       "was specified both as unnamed and named. Named " .
                       "value discarded.";
             $X->olledb_message(-1, 1, 10, $msg);
             next;
          }

          $no_of_pars++;
          $all_parameters[$parno] = \$$named{$key};
       }

       if ($no_of_errs) {
          my $msg = "There were $no_of_errs unknown parameter(s). " .
                    "Cannot execute procedure '$SP'";
          $X->olledb_message(-1, 1, 16, $msg);
          return (wantarray ? () : undef);
       }
    }

    # Compose the SQL statement and initiliaze the batch. We enter the
    # return value as parameter, and start to build the log string.
    my $SP_conv = $X->{'procs'}{$SP}{'normal'};
    $X->do_conversion('to_server', $SP_conv);
    my $sqlstmt = "{? = call $SP_conv";
    if ($no_of_pars > 0) {
       $sqlstmt .= '(' .join(',', ('?') x $no_of_pars) . ')';
    }
    $sqlstmt .= '}';
    $X->initbatch($sqlstmt);
    $X->{ErrInfo}{SP_call} = "EXEC $SP_conv ";

    # Loop over all parameter references  to enter them.
    foreach my $par_ix (0..$#all_parameters) {
       next if not defined($all_parameters[$par_ix]);

       my($param, $is_ref, $value, $name, $maxlen, $precision, $scale, $type,
          $is_input, $is_output, $typeinfo);

       # Get the actual parameter. What is in @all_parameter is a reference to
       # the parameter.
       $param = ${$all_parameters[$par_ix]};

       # And to confuse you even more - the parameter can itself be a reference
       # to the value. (And damn it! The value can also be a reference!)
       $is_ref = (ref $param) =~ /^(SCALAR|REF)$/;

       # Get attributes for the parameters.
       $name      = $$paramdefs[$par_ix]{'name'};
       $type      = $$paramdefs[$par_ix]{'type'};
       $is_output = $$paramdefs[$par_ix]{'is_output'};
       $is_input  = $$paramdefs[$par_ix]{'is_input'};
       $maxlen    = $$paramdefs[$par_ix]{'max_length'};
       $precision = $$paramdefs[$par_ix]{'precision'};
       $scale     = $$paramdefs[$par_ix]{'scale'};
       $typeinfo  = $$paramdefs[$par_ix]{'typeinfo'};

       # Save reference where to receive the < of output parameters.
       if ($is_output) {
          if ($is_ref) {
             push(@output_params, $param);
          }
          else {
             push(@output_params, $all_parameters[$par_ix]);
             if ($^W and not $X->{ErrInfo}{NoWhine}) {
                my $msg = "Output parameter '$name' was not passed as reference";
                $X->olledb_message(-1, 1, 10, $msg);
             }
          }
       }

       # Get the value and perform conversions of name and value.
       $value = ($is_ref ? $$param : $param) if $is_input;
       if (defined $value) {
          $X->do_conversion('to_server', $value);
       }
       $X->do_conversion('to_server', $name);

       # SQL Server 6.x thinks an empty string and NULL is the same, so
       # pass an empty string as one space to 6.5.
       if ($X->{SQL_version} =~ /^6\./ and defined $value and
           length($value) == 0 and $STRINGTYPES{$type}) {
          $value = " ";
       }

       # Set max length for some types where the query does not give the best
       # fit.
       if ($LARGETYPES{$type}) {
          $maxlen = -1;
       }
       elsif ($UNICODETYPES{$type} and $maxlen > 0) {
          $maxlen = $maxlen / 2;
       }

       # Precision and scale should be set only for decimal types.
       if (not $DECIMALTYPES{$type}) {
          $precision = 0;
          $scale     = 0;
       }

       # Add to the log string, execept for return values.
       if ($is_input) {
          $X->{ErrInfo}{SP_call} .= $name . ' = ' .
                                   $X->valuestring($type, $value) . ', ';
       }

       # Now we can enter the parameter.
       $X->enterparameter($type, $maxlen, $name, $is_input, $is_output,
                          $value, $precision, $scale, $typeinfo);
    }

    # Do logging.
    $X->{ErrInfo}{SP_call} =~ s/,\s*$//;
    $X->do_logging;

    # Some variables that we need to execute the function and retrieve the
    # result set.
    my($exec_ok, @results, $resultref);

    # Execute the procedure, unless NoExec is in effect.
    unless ($X->{'NoExec'}) {
       $exec_ok = $X->executebatch();
    }
    else {
       $X->cancelbatch;
       $exec_ok = 0;
    }

    # Retrieve the result sets.
    if (wantarray) {
       @results = $X->do_result_sets($exec_ok, $rowstyle, $resultstyle, $keys);
    }
    else {
       $resultref = $X->do_result_sets($exec_ok, $rowstyle, $resultstyle, $keys);
    }

    # Retrieve output parameters. They are not available if command was
    # cancelled or some such.
    if ($X->getcmdstate == CMDSTATE_GETPARAMS) {
       my ($output_from_sp);

       # Retrieve output parameters
       $X->getoutputparams(undef, $output_from_sp);
       $X->do_conversion('to_client', $output_from_sp);

       # And map values to the input parameters.
       foreach my $ix (0..$#output_params) {
          ${$output_params[$ix]} = $$output_from_sp[$ix];
       }

       # Check the return status if there was one. (The return value is
       # $$retvalueref now.)
       if ($$paramdefs[0]{'is_retstatus'}) {
          my ($retvalue) = $$retvalueref;
          if ($retvalue ne 0 and $X->{ErrInfo}{CheckRetStat} and
              not $X->{ErrInfo}{RetStatOK}{$retvalue}) {
              $X->olle_croak("Stored procedure $SP returned status $retvalue");
          }
       }
    }

    # Remove the faked call from ErrInfo
    delete $X->{ErrInfo}{SP_call};

    # Return the result sets.
    return (wantarray ? @results : $resultref);
}

#-------------------------  sql_insert  -------------------------------
sub sql_insert {
    my($X) = (ref @_[$[] eq PACKAGENAME ? shift @_ : $def_handle);
    my($tblspec) = shift @_;
    my(%values) = %{shift @_};  # Take a copy, we'll be modifying.

    my($tbldef, $col);

    # If have a column profile saved, reuse it.
    if (exists $X->{'tables'}{$tblspec}) {
       $tbldef = $X->{'tables'}{$tblspec};
    }
    else {
       # We don't about this one. Get data about the table from the server.
       my ($objdb, $objid, @columns);

       # Get the object id for the table and it's database
       ($objid, $objdb) = $X->get_object_id($tblspec);
       if (not $objid) {
          my $msg = "Table '$tblspec' is not accessible";
          $X->olledb_message(-1, 1, 16, $msg);
          return;
       }

       # Now, inquire about all the columns in the table and their type.
       # Different handling for different SQL Server versions.
       my $getcols;
       if ($X->{SQL_version} =~ /^6\./) {
          $getcols = <<SQLEND;
              SELECT c.name, type = CASE c.usertype
                                        WHEN 80 THEN ut.name
                                        ELSE t.name
                                    END,
                     "precision" = coalesce(c.prec, 0),
                     scale = coalesce(c.scale, 0), typeinfo = NULL
              FROM   $objdb.dbo.syscolumns c
              JOIN   $objdb.dbo.systypes ut ON c.usertype = ut.usertype
              JOIN   $objdb.dbo.systypes t ON ut.type = t.type
              WHERE  c.id = ?
                AND  t.usertype < 80
                AND  t.name <> 'sysname'
SQLEND
       }
       elsif ($X->{SQL_version} =~ /^[78]\./) {
          $getcols = <<SQLEND;
              SELECT name, type = type_name(xtype), "precision" = prec, scale,
                     typeinfo = NULL
              FROM   $objdb.dbo.syscolumns
              WHERE  id = \@objid
SQLEND
       }
       else {
          # SQL Server 2005 or later.
          $getcols = <<SQLEND;
              SELECT c.name, type = CASE c.system_type_id
                                         WHEN 240 THEN 'UDT'
                                         ELSE type_name(c.system_type_id)
                                     END, c.precision, c.scale,
                     typeinfo =
                     CASE c.system_type_id
                          WHEN 240
                          THEN  coalesce(nullif(\@objdb, ''),
                                         quotename(db_name())) + '.' +
                                quotename(s1.name) + '.' + quotename(t.name)
                          WHEN 241
                          THEN  coalesce(nullif(\@objdb, ''),
                                         quotename(db_name())) + '.' +
                                quotename(s2.name) + '.' + quotename(x.name)
                     END
              FROM   $objdb.sys.all_columns c
              LEFT   JOIN ($objdb.sys.types t
                          JOIN  $objdb.sys.schemas s1 ON t.schema_id = s1.schema_id)
                  ON  c.user_type_id = t.user_type_id
                 AND  t.is_assembly_type = 1
              LEFT   JOIN ($objdb.sys.xml_schema_collections x
                           JOIN  $objdb.sys.schemas s2 ON x.schema_id = s2.schema_id)
                  ON  c.xml_collection_id = x.xml_collection_id
              WHERE  c.object_id = \@objid
SQLEND
       }

       # Trim the SQL from extraneous spaces, to save network bandwidth.
       $getcols =~ s/\s{2,}/ /g;

       # Get the columns. Need special call for 6.5 as we do not support
       # named parameters there.
       if ($X->{SQL_version} =~ /^6\./) {
          $tbldef = $X->internal_sql($getcols, [['int', $objid]],
                                     HASH, KEYED, ['name']);
       }
       else {
          $tbldef = $X->internal_sql($getcols,
                                     {'@objid' => ['int', $objid],
                                      '@objdb' => ['nvarchar', $objdb]},
                                     HASH, KEYED, ['name']);
       }

       # Clear SP_call
       undef $X->{ErrInfo}{SP_call};

       # Save it for future calls.
       $X->{'tables'}{$tblspec} = $tbldef;
    }

    # Build parameter and column array.
    my (@columns, @params);
    foreach my $col (sort keys %values) {
       if (exists $$tbldef{$col}) {
          my $type = $$tbldef{$col}{'type'};
          my $typeinfo = $$tbldef{$col}{'typeinfo'};

          # timestamp columns, cannot be inserted into, so skip.
          next if $type eq 'timestamp';

          if ($DECIMALTYPES{$type}) {
             my $prec = $$tbldef{$col}{'precision'};
             my $scale = $$tbldef{$col}{'scale'};
             $type .= "($prec,$scale)";
          }
          push(@params, [$type, $values{$col}, $typeinfo]);
       }
       else {
          # Missing column an error condition, but let SQL say that.
          push (@params, ['int', undef]);
       }
       if (not defined $values{$col}) {
          $values{$col} = "NULL";
       }
       push(@columns, $col);
    }

    # Build SQL statement.
    my $sqlstmt = "INSERT $tblspec (" . join(', ', @columns) .
                  ")\n   VALUES (" .
                  join(', ', (('?') x scalar(@columns))) . ')';

    # Produce the SQL and run it.
    $X->sql($sqlstmt, \@params);
}

#----------------------- get_result_sets ------------------------------
sub get_result_sets {
   my ($X, $rowstyle, $resultstyle, $keys) = @_;
   check_style_params($rowstyle, $resultstyle, $keys);
   do_result_sets($X, 1, $rowstyle, $resultstyle, $keys);
}

#------------------------- sql_has_errors ----------------------------
sub sql_has_errors {
    my($X) = (ref @_[$[] eq PACKAGENAME ? shift @_ : $def_handle);
    my ($keep) = @_;

    # Check that SaveMessages is on. Warn if not.
    if ($^W and not $X->{ErrInfo}{SaveMessages}) {
       carp "Since ErrInfo.SaveMessages is OFF, it's useless to call sql_has_errors";
    }

    if (not exists $X->{ErrInfo}{Messages}) {
       return 0;
    }

    my $has_error = 0;
    foreach my $msg (@{$X->{ErrInfo}{Messages}}) {
       next unless $msg->{'severity'} >= 11;
       $has_error = 1;
       last;
    }

    if (not $keep and not $has_error) {
       delete $X->{ErrInfo}{Messages};
    }

    return $has_error;
}

#---------------------- sql_get_command_text -------------------------
sub sql_get_command_text {
    my($X) = (ref @_[$[] eq PACKAGENAME ? shift @_ : $def_handle);
    return ($X->{ErrInfo}{SP_call} ? $X->{ErrInfo}{SP_call} :
                                     $X->getcmdtext);
}

#-------------------------  sql_string  -------------------------------
sub sql_string {
    my($X) = (ref @_[$[] eq PACKAGENAME ? shift @_ : $def_handle);
    my($str) = @_;
    if (defined $str) {
       $str =~ s/'/'\'/g;
       "'$str'";
    }
    else {
       "NULL";
    }
}

#------------------------- transaction routines -----------------------
sub sql_begin_trans {
    my($X) = (ref @_[$[] eq PACKAGENAME ? shift @_ : $def_handle);
    $X->sql("BEGIN TRANSACTION");
}

sub sql_commit {
    my($X) = (ref @_[$[] eq PACKAGENAME ? shift @_ : $def_handle);
    $X->sql("COMMIT TRANSACTION");
}

sub sql_rollback {
    my($X) = (ref @_[$[] eq PACKAGENAME ? shift @_ : $def_handle);
    $X->sql("ROLLBACK TRANSACTION");
}

#--------------------- sql_message_handler ----------------------------
sub sql_message_handler {
    my($X, $errno, $state, $severity, $text, $server,
       $procedure, $line, $sqlstate, $source, $n, $no_of_errs) = @_;

    my($ErrInfo, $print_msg, $print_text, $print_lines, $fh);

    # First get a reference to an ErrInfo hash.
    $ErrInfo = $X->{ErrInfo};

    # If this is the first message in a burst, clear the die and carp flags.
    $ErrInfo->{DieFlag}  = 0 if $n == 1;
    $ErrInfo->{CarpFlag} = 0 if $n == 1;

    # Determine where to write the messages.
    $fh = ($ErrInfo->{ErrFileHandle} or \*STDERR);

    # Save messages if requested.
    if ($ErrInfo->{SaveMessages}) {
       my %message;
       tie %message, 'Win32::SqlServer::ErrInfo::Messages';
       %message = (Errno    => $errno,
                   State    => $state,
                   Severity => $severity,
                   Text     => $text,
                   Proc     => $procedure,
                   Line     => $line,
                   Server   => $server,
                   SQLstate => $sqlstate,
                   Source   => $source);
       push(@{$ErrInfo->{Messages}}, \%message);
    }

    # If there is no sqlstate, just set it to empty string, so we don't
    # have to test for undef all the time.
    $sqlstate = '' if not defined $sqlstate;

    # Find out whether we should stop on this error unless die flag
    # already set.
    unless ($ErrInfo->{DieFlag}) {
       if ($severity > $ErrInfo->{MaxSeverity}) {
          $ErrInfo->{DieFlag} = 1 unless ($ErrInfo->{NeverStopOn}{$errno} or
                                          $ErrInfo->{NeverStopOn}{$sqlstate});
       }
       else {
          $ErrInfo->{DieFlag} = ($ErrInfo->{AlwaysStopOn}{$errno} or
                                 $ErrInfo->{AlwaysStopOn}{$sqlstate});
       }
    }

    # Then determine if to print and what.
    unless ($ErrInfo->{NeverPrint}{$errno} or $ErrInfo->{NeverPrint}{$sqlstate}) {
       # Not in neverPrint. If in alwaysPrint, print it all.

       if (not ($ErrInfo->{AlwaysPrint}{$errno} or
                $ErrInfo->{AlwaysPrint}{$sqlstate})) {
          # Nope. Check each part.
          $print_msg = $severity >= $ErrInfo->{PrintMsg};
          $print_text = $severity >= $ErrInfo->{PrintText};
          $print_lines = $severity >= $ErrInfo->{PrintLines};

          # Carp only if there is a message, and severity is above level-
          if ($severity >= $ErrInfo->{CarpLevel} and
              ($print_msg or $print_text or $print_lines)) {
             $ErrInfo->{CarpFlag}++
          }
       }
       else {
          $print_msg = $print_text = $print_lines = 1;
          $ErrInfo->{CarpFlag}++;
       }

       # Here goes printing for each part. First message info.
       if ($print_msg) {
          if (not $source) {
             print $fh "SQL Server message $errno, Severity $severity, ",
                       "State $state";
             print $fh ", Server $server" if $server;
             if ($procedure) {
                print $fh "\nProcedure $procedure, Line $line";
             }
             else {
                print $fh "\nLine $line" if $line;
             }
             print $fh "\n";
          }
          else {
             print $fh "Message "  . ($sqlstate ? $sqlstate : $errno) .
                       " from '$source', Severity: $severity\n";
             print $fh "Internal Win32::SqlServer call: $procedure\n" if $procedure;
          }
       }

       # The text.
       if ($print_text) {
          print $fh "$text\n" if $text;
       }

       # The lines. This is slightly more tricky. If SP_call is defined, use
       # that, else get the command text. Apply LinesWindow only in the latter
       # case.
       if ($print_lines) {
          my ($linetxt, $window);
          $linetxt = $X->sql_get_command_text();
          $window  = $ErrInfo->{LinesWindow};
          if ($linetxt) {
             my ($lineno);
             foreach my $row (split (/\n/, $linetxt)) {
                $lineno++;
                # Always print the line if there is no window or there was
                # no line number. Else print only if lineno is within window.
                if (not defined $window or not $line or
                    $lineno >= $line - $window and $lineno <= $line + $window) {
                   print $fh sprintf("%5d", $lineno), "> $row\n";
                }
             }
          }
       }
    }

    # Check for disconnect. The test on severity is hard-coded as that is
    # how SQL Server works.
    if ($severity >= 20 or $ErrInfo->{DisconnectOn}{$errno} or
       $$ErrInfo{DisconnectOn}{$sqlstate}) {
       $X->disconnect();
    }

    if ($n == $no_of_errs and $ErrInfo->{DieFlag}) {
         $X->olle_croak("Terminating on fatal error");
    }

    if ($n == $no_of_errs and $ErrInfo->{CarpFlag}) {
       carp "Message from " . (defined $source ? $source : 'SQL Server');
    }

    return 1;
}

#---------------------  internal_sql  --------------------------------------
# Very similar to the official sql, but does not check NoExec and Loghandle.
# Use for internal calls to support sql_sp and sql_insert.
sub internal_sql
{
    my($X) = (ref @_[$[] eq PACKAGENAME ? shift @_ : $def_handle);

    my $sql = shift @_;

    # Get parameter array if any.
    my ($arrayparams, $hashparams);
    if (ref $_[0] eq "ARRAY") {
       $arrayparams = shift @_;
    }
    if (ref $_[0] eq "HASH") {
       $hashparams = shift @_;
    }

    # Style parameters. Get them from @_ and then check that values are
    # legal and supply defaults as needed.
    my($rowstyle, $resultstyle, $keys) = @_;
    check_style_params($rowstyle, $resultstyle, $keys);

    # Apply conversion.
    $X->do_conversion('to_server', $sql);

    # Set up the SQL command - initbatch and enter parameters if necesary.
    $X->setup_sqlcmd($sql, $arrayparams, $hashparams);

    my $exec_ok = $X->executebatch;

    # And get the resultsets.
    return $X->do_result_sets($exec_ok, $rowstyle, $resultstyle, $keys);
}

#----------------------- olle_croak, internal -----------------------
sub olle_croak  {
    my ($X, $msg) = @_;
    $X->cancelbatch;
    croak($msg);
}

#---------------------- valuestring, internal----------------------------
sub valuestring {
    my ($X, $datatype, $value) = @_;
    # Returns $value as stringliteral suitable for SQL code.

    if (not defined $value) {
       return "NULL";
    }
    elsif ($UNICODETYPES{$datatype} or $datatype eq 'sql_variant') {
       return 'N' . sql_string($value);
    }
    elsif ($BINARYTYPES{$datatype}) {
       my $ret;
       if ($X->{BinaryAsStr}) {
          $ret = $value;
          $ret = "0x$ret" unless $ret =~ /^0x/i;
       }
       else {
          $ret = "0x" . uc(unpack('H*', $value));
       }
       $ret .= '00' if ($ret eq '0x' and $X->{SQL_version} =~ /^6\./);
       return $ret;
    }
    elsif ($QUOTEDTYPES{$datatype}) {
       return sql_string($value);
    }
    elsif ($datatype eq 'xml') {
       # For xml we need to check the encoding to find out whether we should
       # have an N or not.
       my $encoding;
       my $N = '';
       if ($value =~ /^\<\?xml\s+version\s*=\s*"1.0"\s+encoding\s*=\s*"([^\"]+)"/) {
          $encoding = lc($1);
       }
       if (not $encoding or $encoding =~ /^(utf-16|ucs)/) {
       # If no encoding found, it is UTF-8. If no listed encoding, it is
       # assumed to be 8-bit (or more exactly varchar.)
           $N = 'N';
       }
       elsif ($encoding eq 'utf-8') {
       # An explicit utf-8 declaration is devilish, because the string
       # we will print will not interpreted as UTF-8 by the T-SQL parser.
       # So to make it execute and pass the test suite - we simply remove
       # the part of the declartion! Then we pretend as if it was ucs-2.
          $value =~ s/(^\<\?xml\s+version\s*=\s*"1.0"\s+)encoding\s*=\s*"utf-8"/$1/i;
          $N = 'N';
       }
       return $N . sql_string($value);
    }
    else {
       return $value;
    }
}

#--------------------- new_err_info, internal----------------------------
sub new_err_info {
    # Initiates an err_info hash and returns a reference to it. We
    # set default to print everything but two messages (changed db
    # and language) and to stop on everything above severity 10.

    my(%ErrInfo);
    tie %ErrInfo, 'Win32::SqlServer::ErrInfo';

    # Initiate default error handling: stop on severity > 10, and print
    # both messages and lines.
    $ErrInfo{PrintMsg}       = 1;
    $ErrInfo{PrintText}      = 0;
    $ErrInfo{PrintLines}     = 11;
    $ErrInfo{NeverPrint}     = {'5701' => 1, '5703' => 1};
    $ErrInfo{AlwaysPrint}    = {'3606' => 1, '3607' => 1, '3622' => 1};
    $ErrInfo{MaxSeverity}    = 10;
    $ErrInfo{CheckRetStat}   = 1;
    $ErrInfo{SaveMessages}   = 0;
    $ErrInfo{CarpLevel}      = 10;
    $ErrInfo{DisconnectOn}   = {'2745'  => 1,  '4003' => 1,  '5702' => 1,
                                '17308' => 1, '17310' => 1, '17311' => 1,
                                '17571' => 1, '18002' => 1, '08001' => 1,
                                '08003' => 1, '08004' => 1, '08007' => 1,
                                '08S01' => 1};

    \%ErrInfo;
}

#----------------------- get_codepage_from_reg, internal -------------
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

    $regref->QueryValueEx($cp_value, $dummy, $result);
    $regref->Close or warn "Could not close registry key.\n";

    $result;
}

#-------------------- do_conversion, internal ----------------
sub do_conversion{
    my ($X) = shift @_;
    my ($direction) = shift @_;
    if (defined $X->{$direction}) {
       my $reftype = ref $_[0];
       if ($reftype eq "HASH") {
          %{$_[0]} = &{$X->{$direction}}(%{$_[0]});
       }
       elsif ($reftype  eq "ARRAY") {
          &{$X->{$direction}}(@{$_[0]});
       }
       elsif ($reftype eq "SCALAR") {
          &{$X->{$direction}}(${$_[0]});
       }
       else {
          &{$X->{$direction}}(@_);
       }
    }
}

#------------------------ do_logging, internal ----------------------
sub do_logging {
   my($X) = @_;

   if ($X->{LogHandle}) {
      my ($F) = $X->{LogHandle};
      my $sql = $X->sql_get_command_text();
      print $F "$sql\ngo\n";
   }
}

#--------------------- check_style_params, internal -------------------
sub check_style_params {
# Checks that row- and resultstyle parameters are correct, and provides
# defaults.

    # This is how the parameters eventually will be arranged on return.
    my($rowstyleref)    = \$_[0];
    my($resultstyleref) = \$_[1];
    my($keysref)        = \$_[2];

    my ($rowstyle, $resultstyle, $keys);

    my $rowdefault    = HASH;
    my $resultdefault = SINGLESET;

    # The simple case, just the defaults.
    if (not defined $_[0] and not defined $_[1]) {
       $rowstyle    = $rowdefault;
       $resultstyle = $resultdefault;
    }
    elsif (defined $_[0] and grep ($_ == $_[0], ROWSTYLES)) {
    # First parameter is row style. Next must be result style or be undefined.
       $rowstyle = $_[0];
       $resultstyle = $_[1] || $resultdefault;

       unless (grep ($_ == $resultstyle, RESULTSTYLES) or
               ref $resultstyle eq "CODE") {
          croak PACKAGENAME . ": Illegal resultstyle value: $resultstyle";
       }
    }
    elsif (defined $_[1] and grep ($_ == $_[1], ROWSTYLES)) {
    # Second parameter is row style. First must be result style or be undefined.
       # The default.
       $rowstyle = $_[1];
       $resultstyle = $_[0] || $resultdefault;

       unless (grep ($_ == $resultstyle, RESULTSTYLES) or
               ref $resultstyle eq "CODE") {
          croak PACKAGENAME . ": Illegal resultstyle value: $resultstyle";
       }
    }
    elsif (defined $_[0] and
           (grep ($_ == $_[0], RESULTSTYLES) or ref $_[0] eq "CODE")) {
    # First parameter is result style and second is not row style, but may be
    # keys.
       $resultstyle = $_[0];

       # Move keys if there are any set row style to default.
       if (defined $_[1] and ref $_[1] eq 'ARRAY') {
          $keys = $_[1];
          $rowstyle = $rowdefault;
       }
       elsif (not defined $_[1]) {
          $rowstyle = $rowdefault;
       }
       else {
          croak PACKAGENAME . ": Illegal rowstyle value: $_[1]";
       }
    }
    elsif (defined $_[1] and
           (grep ($_ == $_[1], RESULTSTYLES) or ref $_[1] eq "CODE")) {
    # Second parameter is result style and second is not row style.
       $resultstyle = $_[1];

       # First parameter must be undef.
       if (not defined $_[0]) {
          $rowstyle = $rowdefault;
       }
       else {
          croak PACKAGENAME . ": Illegal rowstyle value: $_[0]";
       }
    }
    else {
       $_[0] = '' if not defined $_[0];
       $_[1] = '' if not defined $_[1];
       croak PACKAGENAME . ": Illegal rowstyle and/or resultstyle values: $_[0], $_[1]";
    }

    # Final check that style parameters are legal
    unless (grep ($_ == $rowstyle, ROWSTYLES)) {
       croak PACKAGENAME . ": Illegal rowstyle value: $rowstyle";
    }
    unless (grep ($_ == $resultstyle, RESULTSTYLES) or
            ref $resultstyle eq "CODE") {
       croak PACKAGENAME . ": Illegal resultstyle value: $resultstyle";
    }

    # If result style is KEYED, check that we have a sensible keys.
    if ($resultstyle == KEYED) {
       $keys = $_[2] unless $keys;
       croak PACKAGENAME . ": No keys given for result style KEYED"
             unless $keys;
       croak PACKAGENAME . ": \$keys is not a list reference"
             unless ref $keys eq "ARRAY";
       croak PACKAGENAME . ": Empty key array given for resultstyle KEYED"
             if @$keys == 0;
       if ($rowstyle != HASH) {
          croak PACKAGENAME . ": \@\$keys must be numeric for rowstyle LIST/SCALAR"
             if grep(/\D/, @$keys);
       }
    }

    # And set the in/out parameters
    $$rowstyleref    = $rowstyle;
    $$resultstyleref = $resultstyle;
    $$keysref        = $keys;
}

#------------------- setup_sqlcmd, internal --------------------------
sub setup_sqlcmd {
   my($X, $sql, $arrayparams, $hashparams) = @_;
   # Common routine for sql and sql_one. If both $params parameters are
   # undef, just calls initbatch. Else runs through the parameters and
   # Generates a call to sp_executesql for $sql, the parameter list and
   # the parameters in %$params. (With a twist for SQL 6.5.)

   # Initial cleanup.
   delete $X->{ErrInfo}{SP_call};

   if (not ($arrayparams or $hashparams)) {
      # This is the simple one. Do it and leave.
      $X->initbatch($sql);
      return 1;
   }

   my (@paramnames);    # A parallel array to $arrayparams that holds the parameter names.
   my ($no_of_unnamed); # The number of elements initially in @$arrayparams.
   my ($paramdecls);    # Parameter declaration for the second param to sp_executesql.
   my (@parameters);    # Here we assemble input to enterparameter.
   my ($paramvalues);   # Parameter assignments for sp_executesql.

   # Named parameters not supported for SQL 6.5
   if ($X->{SQL_version} =~ /^6\./ and $hashparams) {
      $X->olle_croak("Cannot use named parameters with SQL 6.5");
   }

   # Give the all array parameters names on the form @P1 etc
   foreach my $ix (0..$#$arrayparams) {
      my $parno = $ix + 1;
      push(@paramnames, "\@P$parno");
   }

   # Repack hash parameters as array parameters, so we can handle them in
   # the same manner. Also check for name clashes with unnamed parameters.
   $no_of_unnamed = scalar(@$arrayparams);
   foreach my $parname (sort keys %$hashparams) {
      my $parname_as_given = $parname;

      # If the parameter does not have a leading @, add one, and check for
      # clashes.
      if ($parname !~ /^\@/) {
         if (exists $$hashparams{'@' . $parname}) {
            my $msg = "Warning: hash parameters for Win32::SqlServer::sql " .
                      "includes the key '$parname' as well as '\@$parname'. The " .
                      "value for the key '$parname' is discarded.";
            $X->olledb_message(-1, 1, 10, $msg);
            next;
         }
         $parname = '@' . $parname;
      }

      # If name is @P1 or simlar, checj for clash with named parameter.
      if ($parname =~ /^\@P(\d+)$/) {
         my $parno = $1;
         if ($parno <= $no_of_unnamed and $^W) {
            my $msg =  "Warning: Value was provided for a named parameter " .
                       "'\@P$parno', but $no_of_unnamed unnamed values were " .
                       "also provided. The value for the named parameter is " .
                       "discarded.";
            $X->olledb_message(-1, 1, 10, $msg);
            next;
         }
      }

      push(@$arrayparams, $$hashparams{$parname_as_given});
      push(@paramnames, $parname);
   }

   # Now we can iterate over all parameters.
   foreach my $ix (0..$#$arrayparams) {
      my ($par, $parname, $value, $datatype, $isoutput, $length, $precision,
          $scale, $dtypstr, $typeinfo);

      $par = $$arrayparams[$ix];
      $parname = $paramnames[$ix];
      if (ref $par eq 'ARRAY') {
         $datatype  = $$par[0];
         $value     = $$par[1];
         $typeinfo  = $$par[2];
      }
      else {
         $value = $par;
      }

      # If there is no datatype, supply a default, but give a warning unless
      # a NULL value is being passed.
      if (not defined $datatype) {
         if (defined $value and $^W) {
            my $msg = "Warning: no datatype provided for parameter '$parname', value '$value'.";
            $X->olledb_message(-1, 1, 10, $msg);
         }
         $datatype = 'char';
      }

      # Normalize $datatype to be lowercase, except UDT that should be
      # uppercase. Only the part before the first paren should be handled.
      $datatype =~ s/^(\w+)(\(|$)/\L$1\E$2/g;
      $datatype =~ s/^(udt)(\(|$)/\U$1\E$2/g;

      # if datatype includes length or prec/scale, extract it.
      if ($datatype =~ /\(([^\)]+)\)\s*$/) {
         $dtypstr = $datatype;
         my $paren = $1;

         # Save the datatype name temporarily, stripped from the paren stuff.
         # Once we know that the paren value is OK, we save the type in
         # $datatype. This is so that if things are not OK, enterparameter
         # will invoke the message handler, because we should not croak on
         # on this error here.
         my $dtyptemp = $datatype;
         $dtyptemp =~ s/\s*\(.*$//g;

         if ($paren =~ /^\s*(\d+)\s*$/) {
            # One number in parens. This is OK for strings, binary and
            # decimal types
            if ($VARLENTYPES{$dtyptemp}) {
               $length = $1;
               $datatype = $dtyptemp;
            }
            elsif ($DECIMALTYPES{$dtyptemp}) {
               $precision = $1;
               $datatype = $dtyptemp;
            }
         }
         elsif ($paren =~ /^\s*MAX\s*$/i) {
            # MAX, OK for strings and binary.
            if ($MAXTYPES{$dtyptemp}) {
               $dtypstr = "$dtyptemp(MAX)";
               $datatype = $dtyptemp;
               $length = -1;
            }
         }
         elsif ($paren =~ /^\s*(\d+)\s*,\s*(\d+)\s*$/ and
                $DECIMALTYPES{$dtyptemp}) {
             $precision = $1;
             $scale     = $2;
             $datatype = $dtyptemp;
         }
         elsif ($TYPEINFOTYPES{$dtyptemp}) {
             if (defined $typeinfo and $typeinfo ne $paren) {
                my $msg = "Conflicting type information ('$paren' and " .
                          "'$typeinfo') provided for parameter '$parname' " .
                          "of datatype $dtyptemp.";
                $X->olledb_message(-1, 1, 16, $msg);
                return 0;
             }
             $datatype = $dtyptemp;
             $typeinfo = $paren;
         }
      }

      # Get length for variable length types.
      if ($VARLENTYPES{$datatype} and defined $value) {
         unless (defined $length) {
            # Length is always at least 1.
            $length = (length($value) or 1);

            # For binary as string, length passed is only half of value.
            if ($BINARYTYPES{$datatype} and $X->{BinaryAsStr}) {
               $length -= 2 if $value =~ /^0x/ and $length > 2;
               $length++ if $length % 2;   # Make sure it's an even number.
               $length = $length / 2;
            }

            # Handle overlong strings.
            my $maxlen = ($X->{SQL_version} =~ /^6\./ ? 255  :
                         ($UNICODETYPES{$datatype}    ? 4000 :
                                                        8000));
            if ($length > $maxlen) {
               if ($X->{SQL_version} =~ /^[678]\./) {
                  $length = $maxlen;
               }
               else {
                  # On SQL 2005 and later we can use MAX for some datatypes
                  $length = ($MAXTYPES{$datatype} ? -1 : $maxlen);
               }
            }
         }

         # Form the data type string.
         unless (defined $dtypstr) {
            $dtypstr = $datatype . '(' . (($length >= 0) ? $length : 'MAX') . ')';
         }
      }
      elsif ($LARGETYPES{$datatype}) {
         $length = -1;
      }
      else {
         $length = 0;
      }

      # Set precision for decimal types if not provided.
      if ($DECIMALTYPES{$datatype}) {
         if (not defined $precision or not defined $scale) {
            if ($^W and defined $value) {
               my $msg = "Precision and/or scale missing for decimal parameter '$parname'.";
               $X->olledb_message(-1, 1, 10, $msg);
            }
            $precision = 18 if not defined $precision;
            $scale     = 0  if not defined $scale;
         }
         unless (defined $dtypstr) {
            $dtypstr = "$datatype($precision, $scale)";
         }
      }
      else {
         $precision = 0;
         $scale     = 0;
      }

      # Check that typeinfo not provided when not applicable, and that is
      # specified for UDT.
      if ($TYPEINFOTYPES{$datatype}) {
         if ($datatype eq 'UDT' and not defined $typeinfo) {
            my $msg = "No actual user type specified for UDT parameter '$parname'.";
            $X->olledb_message(-1, 1, 16, $msg);
            $X->cancelbatch;
            return 0;
         }

         if ($datatype eq 'UDT') {
             $dtypstr  = $typeinfo;
         }
         elsif ($datatype eq 'xml') {
             $dtypstr = $datatype . ($typeinfo ? "($typeinfo)" : "");
         }
      }
      elsif (defined $typeinfo) {
         undef $typeinfo;
      }

      unless (defined $dtypstr) {
         $dtypstr = $datatype;
      }

      # Do conversion of value and parameter name
      $X->do_conversion('to_server', $value);
      $X->do_conversion('to_server', $parname);

      # And save the parameter.
      push(@parameters, [$datatype, $length, $parname, 1, 0, $value,
                         $precision, $scale, $typeinfo]);

      # Add to the parameter declaration.
      $paramdecls .= (defined $paramdecls ? ", " : '') .
                     $parname . " " . $dtypstr;

      # Add to the parameter string for logging.
      $paramvalues .= (defined $paramvalues ? ", " : '') .
                       $parname . " = " . $X->valuestring($datatype, $value);
   }

   unless ($X->{SQL_version} =~ /^6\./) {
      # Replace ? with @P1 etc in the query string.
      $X->replaceparamholders($sql);

      # Build log string for error handling.
      $X->{errInfo}{SP_call} = "EXEC sp_executesql N" . sql_string($sql) . ",\n" .
                               ' ' x 5 . 'N'. sql_string($paramdecls) . ",\n" .
                               ' ' x 5 . $paramvalues;

      # First build the sp_executesql command and init the batch, and enter
      # the first parameter.
      my $executesql = '{call sp_executesql(?, ?, ' .
                     join(', ', ('?') x scalar(@parameters)) . ')}';
      $X->initbatch($executesql);

      # Enter parameter for the statement. On SQL 2005, we can use
      # nvarchar(max), but not on SQL7/2000.
      my $len = ($X->{SQL_version} =~ /^[78]\./ ? length($sql) : -1);
      $X->enterparameter('nvarchar', $len, '@stmt', 1, 0, $sql);

      # Enter the parameter for parameter list.
      $len = ($X->{SQL_version} =~ /^[78]\./ ? length($paramdecls) : -1);
      $X->enterparameter('nvarchar', $len, '@parameters', 1, 0, $paramdecls);
   }
   else {
      # On 6.5, sp_executesql is not available, so we don't replace the
      # holders, but leave this to SQLOLEDB.
      $X->initbatch($sql);

      # And supplement SP_call with the real SQL statement and put the
      # sp_executesql thing as a comment.
      $X->{errInfo}{SP_call} = $sql . "\n/*" .
                               ' ' x 3 . 'N'. sql_string($paramdecls) . ",\n" .
                               ' ' x 5 . $paramvalues . ' */';
   }

   # Enter all the "real" parameters.
   foreach my $p (@parameters) {
      $X->enterparameter(@$p);
   }

   return 1;
}

#---------------------- get_sqlserver_version -------------------------
# Retieves the SQL Server version. Since this routine may be called from
# FETCH, we have to tread carefully, and not call code were may happen to
# look at SQL_version!
sub get_sqlserver_version {
    my($self) = @_;

    my ($exec_ok, $sqlver);

    $self->initbatch("EXEC master.dbo.xp_msver 'ProductVersion'");
    $exec_ok = $self->executebatch();
    $self->olle_croak("Could not retrieve SQL Server version.\n")
        if not $exec_ok;
    while ($self->nextresultset()) {
       my $hashref;
       while ($self->nextrow($hashref, undef)) {
         $sqlver = $$hashref{'Character_Value'};
         last if $sqlver;
       }
       last if $sqlver;
    }
    if (not $sqlver) {
       $self->olle_croak("Could not retrieve SQL Server version.\n");
    }
    $self->cancelbatch();
    return $sqlver;
}

#------------------- get_object_id, internal ---------------------------
sub get_object_id {
   my($X, $objspec) = @_;
# Retrieves the object id for a database object.

    my(@objspec, $server, $objdb, $schema, $object, $objid, $normalspec);

    # Call C++ code to crack the object specification into parts.
    $X->parsename($objspec, 1, $server, $objdb, $schema, $object);

    # If we get a server, we are not even going to try. This cannot work.
    return (undef, undef) if $server;

    # Construct a normalised object specification. This is basically the
    # input, but spaces between the parts removed.
    $normalspec = ($objdb ? "$objdb." : '') .
                  (($schema or $objdb) ? "$schema." : '') .
                   $object;

    # A temporary object is per definition in tempdb.
    if ($object =~ /^#/ and $objdb eq '') {
       $objdb = "tempdb";
    }

    # Now we can reconstruct the object specification.
    $objspec = "$objdb.$schema.$object";

    # Get the object-id.
    $objid = $X->internal_sql("SELECT object_id(?)", [['nvarchar', $objspec]],
                              SCALAR, SINGLEROW);

    # Here is a gotcha on SQL 6.5 "db..sp_help" actually gives an object id,
    # despite the SP being in master, so we must double-check.
    if ($X->{SQL_version} =~ /^6\./ and $objdb ne '' and
        $object =~ /^[\"\[]?sp_/) {
        my $sql = "SELECT id FROM $objdb..sysobjects WHERE name = ?";
        $objid = $X->internal_sql($sql, [['nvarchar', $object]],
                                  SCALAR, SINGLEROW);
    }

    # If no luck, it might still be a system procedure.
    if (not defined $objid and $object =~ /^[\"\[]?sp_/) {
       $objdb = "master";
       $objspec = "master.$schema.$object";
       $objid = $X->internal_sql("SELECT object_id(?)", [['nvarchar', $objspec]],
                                 SCALAR, SINGLEROW);
    }

    # Clear SP_call from error info to avoid incorrect statement prints.
    undef $X->{ErrInfo}{SP_call};

    # Return id, database and normalised spec.
    ($objid, $objdb, $normalspec);
}

#---------------------- do_result_sets, internal ---------------------------------
sub do_result_sets {
    my($X, $exec_ok, $rowstyle, $resultstyle, $keys) = @_;

    my ($morerows, $userstat, $is_callback, $isregular, $ix, $ressetno, $dataref,
        $resref, $keyed_res, $iscancelled, $caller);

    $is_callback = ref $resultstyle eq "CODE";
    $isregular   = grep ($_ == $resultstyle, (MULTISET, SINGLESET, SINGLEROW));
    $iscancelled = not $exec_ok;

    $ix = $ressetno = 0;
    $userstat = RETURN_NEXTROW;
    while (not $iscancelled and $X->isconnected() and $X->nextresultset) {
       $ressetno++;

       # He said NORESULT? Cancel the query, and proceed to next.
       if ($resultstyle == NORESULT) {
          $X->cancelresultset;
          next;
       }

       # For the regular result styles create an empty array, if there is none at
       # the current index.
       if ($isregular) {
          @{$$resref[$ix]} = () unless defined @{$$resref[$ix]};
       }
       elsif ($resultstyle == KEYED) {
          # For KEYED create result set, now we know we have a result set.
          $keyed_res = {} unless $keyed_res;
       }

       do {
          $morerows = $X->nextrow(($rowstyle == HASH) ? $dataref : undef,
                                  ($rowstyle == HASH) ? undef : $dataref);

          if ($morerows) {
             # Convert to client charset before anything else.
             $X->do_conversion('to_client', $dataref);

             # For SCALAR convert to joined string. (But for KEYED, this is deferred.)
             if ($rowstyle == SCALAR and $resultstyle != KEYED) {
                $dataref = list_to_scalar($dataref);
             }

             # Save the row if we have a regular resultstyle.
             if ($isregular) {
                push(@{$$resref[$ix]}, $dataref);
             }
             elsif ($resultstyle == KEYED) {
                # This is keyed access.
                store_keyed_result($X, $rowstyle, $keys, $dataref, $keyed_res);
             }
             elsif ($is_callback) {
                $userstat = &$resultstyle($dataref, $ressetno);

                if ($userstat == RETURN_NEXTQUERY) {
                   # He wants next result set, so leave this one.
                   $X->cancelresultset;
                   $morerows = 0;
                }
                elsif ($userstat != RETURN_NEXTROW) {
                # Whatever, cancel the entire batch.
                   $morerows = 0;
                   $iscancelled = 1;
                   $X->cancelbatch;
                   if ($userstat == RETURN_ABORT) {
                      $X->olle_croak("User-supplied callback returned RETURN_ABORT");
                   }
                   elsif ($userstat != RETURN_CANCEL and $userstat != RETURN_ERROR) {
                      $X->olle_croak("User-supplied callback returned unknown return code");
                   }
                }
             }
          }
       }  until not $morerows;

       # If multiset requested advance index
       $ix++ if $resultstyle == MULTISET;
    }

    if ($is_callback) {
       return $userstat;
    }
    elsif (wantarray) {
       if ($resultstyle == KEYED) {
          if (defined $keyed_res) {
             return %$keyed_res;
          }
          else {
             return ();
          }
       }
       elsif (defined $resref) {
          if    ($resultstyle == MULTISET)  {return @$resref }
          elsif ($resultstyle == SINGLESET) {return @{$$resref[0]} }
          elsif ($resultstyle == SINGLEROW) {
              if    ($rowstyle == HASH)
                 { return (defined $$resref[0][0] ? %{$$resref[0][0]} : () )}
              elsif ($rowstyle == LIST)
                 { return (defined $$resref[0][0] ? @{$$resref[0][0]} : () )}
              elsif ($rowstyle == SCALAR) { return @{$$resref[0]} }
          }
          elsif ($resultstyle == KEYED) { return %$keyed_res; }
          else  { return ()}
       }
       else {
          return ();
       }
    }
    else {
       if    ($resultstyle == MULTISET)  {return $resref }
       elsif ($resultstyle == SINGLESET) {return $$resref[0] }
       elsif ($resultstyle == SINGLEROW) {return $$resref[0][0] }
       elsif ($resultstyle == KEYED)     {return $keyed_res }
       else  { return undef}
    }
}

#----------------------------- list_to_scalar ------------------------
# This routine takes a data array and returns a scalar from it. Care
# if being taken to avoid "unitialized value" warnings.
sub list_to_scalar {
   my ($arr) = @_;
   local($^W) = 0;
   if (@$arr == 0) {
      return undef;
   }
   elsif (@$arr == 1) {
      # If there is a single element return this as is and do not use
      # join below, as this would convert an undef to defined value.
      return $$arr[0];
   }
   else
   {
      return join($SQLSEP, @$arr);
   }
}


#------------------------------ store_keyed_result ---------------------
# This routine implements KEYED access. The key columns are removed from the
# list/hash that $dataref points to and added as keys to $keyed_res.
sub store_keyed_result {
   my ($X, $rowstyle, $keys, $dataref, $keyed_res) = @_;

   my ($keyvalue, $keyname, $keyno, $ref, $keystr);

   $ref = $keyed_res;
   $keystr = "";

   # Loop over the keys.
   foreach my $ix (0..$#$keys) {
      # First find the key value, different strategies with different row styles.
      if ($rowstyle == HASH) {
         # Get the key name.
         $keyname = $$keys[$ix];

         # If the key does not exist, we give up.
         unless (exists $$dataref{$keyname}) {
            $X->olle_croak(PACKAGENAME . ": No key '$keyname' in result set");
         }

         # Get the key value, and delete it from the data.
         $keyvalue = $$dataref{$keyname};
         delete $$dataref{$keyname};
      }
      else {
         # Now we have a key number.
         $keyno = $$keys[$ix];

         # It must be a valid index in the result set.
         unless ($keyno >= 1 and $keyno <= $#$dataref + 1) {
             $X->olle_croak(PACKAGENAME . ": Key number '$keyno' is not valid in result set");
         }

         # Get the key value, but don't touch @$dataref yet.
         $keyvalue = $$dataref[$keyno - 1];
      }

      # If this is not the last key, just create the node.
      if ($ix < $#$keys) {
         $ref = \%{$$ref{$keyvalue}};
      }

      # Add keys to debug string, for use in warning messages.
      $keystr .= "<$keyvalue>" if $^W;
   }

   # Now we can remove data from an array - had we done this above, the key numbers
   # wouldn't have matched.
   if ($rowstyle != HASH) {
      foreach my $ix (reverse sort @$keys) {
         splice(@$dataref, $ix - 1, 1);
      }

      # If we're talking scalar, convert at this point
      if ($rowstyle == SCALAR) {
         $dataref = list_to_scalar($dataref);
      }
   }


   # At this point $ref{$keyvalue} is where we want to store the rest of the data.
   # Just check that the spot is not already occupied.
   if ($^W) {
      carp "Key(s) $keystr is not unique" if exists $$ref{$keyvalue};
   }

   # And write into the result set.
   $$ref{$keyvalue} = $dataref;
}

package Win32::SqlServer::ErrInfo;

use strict;
use Tie::Hash;
use Carp;

use vars qw(@ISA @EXPORT);

@ISA = qw(Exporter Tie::StdHash);

use constant FIELDS => qw(ErrFileHandle DieFlag CarpFlag MaxSeverity
                          NeverStopOn AlwaysStopOn PrintMsg PrintText
                          PrintLines NeverPrint AlwaysPrint CarpLevel
                          CheckRetStat RetStatOK SaveMessages Messages
                          SP_call NoWhine LinesWindow DisconnectOn);

my %fields;

foreach my $f (FIELDS) {
   $fields{$f}++;
}


# My own FETCH routine, chckes that retrieval is of a known elements.
sub FETCH {
   my ($self, $key) = @_;
   if (not exists $fields{$key}) {
       $key =~ s/^./uc($&)/e;
       if (not exists $fields{$key}) {
           croak("Attempt to fetch undefined ErrInfo element '$key'");
       }
   }
   return $self->{$key};
}

# My own STORE routine, barfs if attribute is non-existent.
sub STORE {
   my ($self, $key, $value) = @_;
   if (not exists $fields{$key}) {
       $key =~ s/^./uc($&)/e;
       if (not exists $fields{$key}) {
           croak("Attempt to set undefined ErrInfo element '$key'");
       }
   }
   $self->{$key} = $value;
}

sub DELETE {
   my ($self, $key) = @_;
   if (not exists $fields{$key}) {
       $key =~ s/^./uc($&)/e;
   }
   delete $self->{$key};
}

sub EXISTS {
   my ($self, $key) = @_;
   if (not exists $fields{$key}) {
       $key =~ s/^./uc($&)/e;
   }
   return exists $self->{$key};
}


package Win32::SqlServer::ErrInfo::Messages;

use strict;
use Tie::Hash;
use Carp;

use vars qw(@ISA @EXPORT);

@ISA = qw(Exporter Tie::StdHash);

use constant FIELDS => qw(Errno State Severity Proc Line Server
                          Text SQLstate Source);

my %mfields;

foreach my $f (FIELDS) {
   $mfields{$f}++;
}

# The same FETCH as before. Barf if does not exist, but permit initial
# lowercase.
sub FETCH {
   my ($self, $key) = @_;
   if (not exists $mfields{$key}) {
       $key =~ s/^./uc($&)/e;
       if (not exists $mfields{$key}) {
           croak("Attempt to fetch undefined Message element '$key'");
       }
   }
   return $self->{$key};
}

# My own STORE routine, barfs if attribute is non-existent and permits
# inital lowercase.
sub STORE {
   my ($self, $key, $value) = @_;
   if (not exists $mfields{$key}) {
       $key =~ s/^./uc($&)/e;
       if (not exists $mfields{$key}) {
           croak("Attempt to set undefined Message element '$key'");
       }
   }
   $self->{$key} = $value;
}

sub DELETE {
   my ($self, $key) = @_;
   if (not exists $mfields{$key}) {
       $key =~ s/^./uc($&)/e;
   }
   delete $self->{$key};
}

sub EXISTS {
   my ($self, $key) = @_;
   if (not exists $mfields{$key}) {
       $key =~ s/^./uc($&)/e;
   }
   return exists $self->{$key};
}



1;
