/*---------------------------------------------------------------------
 $Header: /Perl/OlleDB/SqlServer.xs 61    06-04-17 21:52 Sommar $

  Copyright (c) 2004-2006   Erland Sommarskog

  $History: SqlServer.xs $
 * 
 * *****************  Version 61  *****************
 * User: Sommar       Date: 06-04-17   Time: 21:52
 * Updated in $/Perl/OlleDB
 * We are now at version 2.002. Changed how the CLSIDs for the provider
 * are saved. Now saving them static directly, and not saving pointers, as
 * the pointers would point somewhere else when an ASP page ran a second
 * time. Also moved CoInitializeEx so it's only called if there is no
 * data_init_ptr.
 *
 * *****************  Version 60  *****************
 * User: Sommar       Date: 05-11-26   Time: 23:47
 * Updated in $/Perl/OlleDB
 * Renamed the module to Win32::SqlServer and advanced to version 2.001.
 *
 * *****************  Version 59  *****************
 * User: Sommar       Date: 05-11-13   Time: 16:33
 * Updated in $/Perl/OlleDB
  ---------------------------------------------------------------------*/


#define UNICODE
#define DBINITCONSTANTS
#define INITGUID
#define _CRT_SECURE_NO_DEPRECATE

#define _WIN32_DCOM   // Needed for CoInitializeEx

#include <windows.h>
#include <assert.h>
//#include <stddef.h>
#include <cguid.h>
#include <oledb.h>
#include <oledberr.h>
#include <msdasc.h>
#include <msdadc.h>
#include <msdaguid.h>

#define _SQLNCLI_OLEDB
#include <SQLNCLI.h>

// Here we include the Perl stuff.
#if defined(__cplusplus)
extern "C" {
#endif

#include "EXTERN.h"
#include "perl.h"
#include "XSUB.h"

#if defined(__cplusplus)
}
#endif

//#include "win32.h"

#undef FILEDEBUG
#ifdef FILEDEBUG
FILE *dbgfile = NULL;
#endif

#define XS_VERSION "2.002"


// This is stuff for init properties. When the module starts up, we set up a
// static array, and then is read-only.
typedef enum init_propsets
    {not_in_use = -1, oleinit_props = 0, ssinit_props = 1, datasrc_props = 2}
init_propsets;
#define NO_OF_INIT_PROPSETS 3
typedef struct {
   char             name[50];       // Name of prop exposed to user.
   init_propsets    propset_enum;   // In which property set property belongs.
   DBPROPID         property_id;    // ID for property in OLE DB.
   VARTYPE          datatype;       // Datatype of the property.
   VARIANT          default_value;  // Default value for the property.
} init_property;
#define MAX_INIT_PROPERTIES 50
static init_property gbl_init_props[MAX_INIT_PROPERTIES];

// This global holds how many of the SSINIT properties that applies to
// SQLOLEDB - there are some that only applies to SQL Native Client.
static int no_of_sqloledb_ssprops;

// This array holds where each property set starts in gbl_init_props;
static struct {
   int start;
   int no_of_props;
} init_propset_info[NO_OF_INIT_PROPSETS];

typedef enum hash_key_id
{
    HV_internaldata,
    HV_propsdebug,
    HV_autoconnect,
    HV_rowsatatime,
    HV_decimalasstr,
    HV_datetimeoption,
    HV_binaryasstr,
    HV_dateformat,
    HV_msecformat,
    HV_udtformat,
    HV_cmdtimeout,
    HV_msgcallback,
    HV_querynotification,
    HV_SQLversion
} hash_key_id;

static char *hash_keys[] =
   { "internaldata", "PropsDebug", "AutoConnect", "RowsAtATime",
     "DecimalAsStr", "DatetimeOption", "BinaryAsStr", "DateFormat",
     "MsecFormat", "UDTFormat", "CommandTimeout", "MsgHandler",
     "QueryNotification", "SQL_version"};

typedef enum provider_enum {
    provider_default, provider_sqloledb, provider_sqlncli
} provider_enum;

typedef enum cmdstate_enum {
    cmdstate_init, cmdstate_enterexec, cmdstate_nextres, cmdstate_nextrow,
    cmdstate_getparams
} cmdstate_enum;

typedef enum dt_options {
   dto_hash, dto_iso, dto_regional, dto_float, dto_strfmt
} dt_options;

typedef enum bin_options {
   bin_binary, bin_string, bin_string0x
} bin_options;

// Options that affect data processing, extracted to a local struct in
// some routines.
typedef struct {
    int         DecimalAsStr;
    bin_options BinaryAsStr;
    dt_options  DatetimeOption;
    char       *DateFormat;
    char       *MsecFormat;
} formatoptions;

// This struct is a big union that holds the value of a parameter to a
// stored procedure.
typedef struct paramvalue {
   union {
      BOOL        bit;
      BYTE        tinyint;
      SHORT       smallint;
      LONG        intval;
      LONGLONG    bigint;
      FLOAT       real;
      DOUBLE      floatval;
      CY          money;
      WCHAR     * nvarchar;
      char      * varchar;
      GUID        guid;
      DB_NUMERIC  decimal;
      BYTE      * binary;
      DBTIMESTAMP datetime;
      SSVARIANT   sql_variant;
   };
} paramvalue;

// This struct describes a parameter to a stored procedure. The caller
// enters one parameter at a time, and we save the parameter in a linked
// list, which we keep until batch is completed.
typedef struct paramdata {
   DBTYPE            datatype;
   BOOL              isinput;
   BOOL              isoutput;
   BOOL              isnull;
   DBLENGTH          value_len;
   paramvalue        value;
   void            * buffer_ptr;   // Copy of varchar/binary pointer, for simple cleanup.
   BSTR              bstr;         // Ditto widechar data, which goes into a different pool.
   DBBINDING         binding;
   DBPARAMBINDINFO   param_info;
   int               param_props_cnt;
   DBPROP          * param_props;
   paramdata       * next;
} paramdata;

// This struct holds data which is entirely local to the XS module. This
// is mainly pointers to various SQLOLEDB objects.
typedef struct {
    // Data source, if non-NULL we are connected.
    IDBInitialize          * init_ptr;
    IDBCreateSession       * datasrc_ptr;
    BOOL                     isautoconnected;  // If connection was through connect() or not.
    provider_enum            provider;         // SQLOLEDB or SQLNCLI

    // Property sets and properies with initialization properties.
    DBPROPSET                init_propsets[NO_OF_INIT_PROPSETS];
    DBPROP                   init_properties[MAX_INIT_PROPERTIES];

    // A command text, possibly with parameters that are being assembled
    // with initbatch and enterparameter.
    BSTR                     pending_cmd;  // (Parameterised) cmd for which caller is supplying parmeters.
    paramdata              * paramfirst;   // Head of linked list for parameters.
    paramdata              * paramlast;    // Tail of parameter list.
    ULONG                    no_of_params; // Length of parameter list.
    ULONG                    no_of_out_params;  // And how many that are outparams.
    BOOL                     params_available;  // Not set until all result sets are exhausted.

    // SQLOLEDB parameters that are created by executebatch.
    IDBCreateCommand       * session_ptr;
    ICommandText           * cmdtext_ptr;
    ICommandWithParameters * paramcmd_ptr;
    ISSCommandWithParameters * ss_paramcmd_ptr;
    IAccessor              * paramaccess_ptr;

    // Data with and about parameters.
    BOOL                     all_params_OK;
    DBPARAMBINDINFO        * param_info;    // Information about parameters.
    DBBINDING              * param_bindings; // How parameters are bound in param_buffer.
    BYTE                   * param_buffer;  // Buffer for parameter values.
    ULONG                    size_param_buffer;  // Size of that buffer.
    HACCESSOR                param_accessor;
    DBBINDSTATUS           * param_bind_status;

    // Pointers for result sets.
    IMultipleResults       * results_ptr;
    BOOL                     have_resultset; // Whether we have an active result set.
    IRowset                * rowset_ptr;
    IAccessor              * rowaccess_ptr;
    HACCESSOR                row_accessor;

    // We get a number of rows at a time into a buffer (because Gert
    // suggested to, but it does not seem to make a difference.)
    HROW                   * rowbuffer;      // Buffered rows from SQLOLEDB.
    ULONG                    rows_in_buffer; // Size of rowbuffer.
    ULONG                    current_rowno;  // Current row in rowbuffer.

    // Data with and about the columns.
    ULONG                    no_of_cols;     // No of columns in current result set.
    DBCOLUMNINFO           * column_info;
    WCHAR                  * colname_buffer; // Memory area for names in *colunm_info.
    DBBINDING              * col_bindings;
    DBBINDSTATUS           * col_bind_status;
    BYTE                   * data_buffer;    // Data buffer for a single row.
    ULONG                    size_data_buffer;  // Size of that data buffer.
    SV                    ** column_keys;    // We convert column names for hash keys once and save them.
} internaldata;



// Type map that maps name of SQL Server types to OLE DB type indicators.
// This two static arrays that we fill on first load. typenames has all
// the name of the types, and typeindicators has on the same index the
// the indicator to use.
#define SIZE_TYPE_MAP 1000
static char   typenames[SIZE_TYPE_MAP];
static DBTYPE typeindicators[SIZE_TYPE_MAP];

// Global pointer to OLE DB Services. Set once when we intialize, and
// never released.
static IDataInitialize * data_init_ptr    = NULL;

// Global pointer the OLE DB conversion library.
static IDataConvert    * data_convert_ptr = NULL;

// Global pointer to the IMalloc interface. Most of the time when we allocate
// memory, we rely on the Perl methods. However, there are situations when
// we must free memory allocated by SQLOLEDB. Same here, we create once, as
// the COM implementation is touted as thread-safe.
static IMalloc*   OLE_malloc_ptr = NULL;

// Global variables for class ids for the two possible providers.
static CLSID  clsid_sqloledb = CLSID_NULL;
static CLSID  clsid_sqlncli = CLSID_NULL;

//---------------------------------------------------------------------
// General convenience routines.
// The first are for getting a value out/in to the OlleDB hash.
//---------------------------------------------------------------------
static SV **fetch_from_hash (SV* olle_ptr, hash_key_id id) {
   HV * hv;
   hv = (HV *) SvRV(olle_ptr);
   return hv_fetch(hv, hash_keys[id], strlen(hash_keys[id]), FALSE);
}

static void delete_from_hash(SV *olle_ptr, hash_key_id id) {
   HV * hv;
   hv = (HV *) SvRV(olle_ptr);
   hv_delete(hv, hash_keys[id], strlen(hash_keys[id]), G_DISCARD);
}


//---------------------- Options ------------------------------------
// These routines retrieves an attribute each from the OlleDB hash.

static SV * fetch_option(SV * olle_ptr, hash_key_id id) {
// Fetches an option from the hash, and only returns an SV, if there is a
// defined value.
   SV  **svp;
   SV  * retsv = NULL;
   svp = fetch_from_hash(olle_ptr, id);
   if (svp != NULL) {
      SvGETMAGIC(*svp);
      if (SvOK(*svp)) {
         retsv = *svp;
      }
   }
   return retsv;
}

static BOOL OptAutoConnect (SV * olle_ptr) {
   SV *sv;
   BOOL retval = FALSE;
   if (sv = fetch_option(olle_ptr, HV_autoconnect)) {
      retval = SvTRUE(sv);
   }
   return retval;
}

static BOOL OptPropsDebug(SV * olle_ptr) {
   SV * sv;
   BOOL retval = FALSE;
   if (sv = fetch_option(olle_ptr, HV_propsdebug)) {
      retval = SvTRUE(sv);
   }
   return retval;
}

static int OptRowsAtATime(SV * olle_ptr) {
   SV * sv;
   int retval = 100;
   if (sv = fetch_option(olle_ptr, HV_rowsatatime)) {
      retval = SvIV(sv);
      if (retval <= 0) {
          retval = 1;
      }
   }
   return retval;
}

static BOOL OptDecimalAsStr(SV * olle_ptr) {
   SV * sv;
   BOOL retval = FALSE;
   if (sv = fetch_option(olle_ptr, HV_decimalasstr)) {
      retval = SvTRUE(sv);
   }
   return retval;
}

static dt_options OptDatetimeOption(SV * olle_ptr) {
   SV * sv;
   dt_options retval = dto_iso;
   if (sv = fetch_option(olle_ptr, HV_datetimeoption)) {
      retval = (dt_options) SvIV(sv);
   }
   return retval;
}

static bin_options OptBinaryAsStr(SV * olle_ptr) {
   SV * sv;
   bin_options retval = bin_binary;
   if (sv = fetch_option(olle_ptr, HV_binaryasstr)) {
      if (SvTRUE (sv)) {
         char * str = SvPV_nolen(sv);
         retval = bin_string;
         if (strcmp(str, "x") == 0) {
            retval = bin_string0x;
         }
      }
   }
   return retval;
}

static char * OptDateFormat(SV * olle_ptr) {
   SV   * sv;
   char * retval = NULL;
   if (sv = fetch_option(olle_ptr, HV_dateformat)) {
      retval = SvPV_nolen(sv);
   }
   return retval;
}

static char * OptMsecFormat(SV * olle_ptr) {
   SV * sv;
   char * retval = NULL;
   if (sv = fetch_option(olle_ptr, HV_msecformat)) {
      retval = SvPV_nolen(sv);
   }
   return retval;
}


static int OptCommandTimeout(SV * olle_ptr) {
   SV * sv;
   int retval = 0;
   if (sv = fetch_option(olle_ptr, HV_cmdtimeout)) {
      retval = SvIV(sv);
   }
   return retval;
}

static HV* OptQueryNotification(SV * olle_ptr) {
   SV * sv;
   HV * retval = NULL;
   if (sv = fetch_option(olle_ptr, HV_querynotification)) {
      retval = (HV *) SvRV(sv);
   }
   return retval;
}


// Handling of the formatoptions struct.
static formatoptions getformatoptions(SV   * olle_ptr) {
   formatoptions opts;

   opts.DecimalAsStr = OptDecimalAsStr(olle_ptr);
   opts.BinaryAsStr = OptBinaryAsStr(olle_ptr);
   opts.DatetimeOption = OptDatetimeOption(olle_ptr);
   if (opts.DatetimeOption == dto_strfmt) {
       opts.DateFormat = OptDateFormat(olle_ptr);
       opts.MsecFormat = OptMsecFormat(olle_ptr);
   }
   else {
       opts.DateFormat = NULL;
       opts.MsecFormat = NULL;
   }

   return opts;
}


// Convenience routine that retrieves the internaldata pointer from the
// internal hash.
static internaldata *get_internaldata(SV *olle_ptr)
{
    HV *hv;
    SV **svp;
    internaldata *ptr;

    if(!SvROK(olle_ptr))
        croak("olle_ptr parameter is not a reference!");
    hv = (HV *)SvRV(olle_ptr);
    if(! (svp = fetch_from_hash(olle_ptr, HV_internaldata)) )
        croak("Internal error: no internaldata key in hash");
    ptr = (internaldata *) SvIV(*svp);
    return ptr;
}

//-------------------------------------------------------------------
// Convenience routines for converting from/to SV to C++ strings, char
// and BSTR. More conversion routines later.
//------------------------------------------------------------------
//Converts a plain ANSI string to a BSTR in Unicode, using SysAllocStr
static BSTR SV_to_BSTR (SV       * sv,
                        DBLENGTH * bytelen = NULL,
                        BOOL       add_BOM = FALSE)
{  int      widelen;
   int      ret;
   DWORD    err;
   BSTR     bstr;
   WCHAR  * tmp;
   STRLEN   sv_len;
   char   * sv_text = (char *) SvPV(sv, sv_len);
   UINT     decoding = (SvUTF8(sv) ? CP_UTF8 : CP_ACP);
   DWORD    flags = (decoding == CP_UTF8 ? 0 : MB_PRECOMPOSED);

   //warn("str = '%s', len = %d, utfblen = %d, utf8 = %x.\n",
   //     sv_text, sv_len, sv_len_utf8(sv), SvUTF8(sv));

   if (sv_len > 0) {
      // First find out how long the wide string will be, by calling
      // MultiByteToWideChar without a buffer.
      widelen = MultiByteToWideChar(decoding, flags, sv_text, sv_len, NULL, 0);

      // Any BOM requires space.
      if (add_BOM) {
         widelen++;
      }

      // Allocate string.
      bstr = SysAllocStringLen(NULL, widelen);

      // Add BOM if required, add move point where to write the converted
      // data one step ahead.
      if (add_BOM) {
         bstr[0] = 0xFEFF;
         tmp = bstr + 1;
      }
      else {
         tmp = bstr;
      }

      // And now for the real thing.
      ret = MultiByteToWideChar(decoding, flags, sv_text, sv_len, tmp, widelen);

      if (! ret) {
         err = GetLastError();
         croak("sv_to_bstr failed with %ld when converting string '%s' to Unicode",
                err, sv_text);
      }
   }
   else {
      bstr = SysAllocString(L"");
      widelen = 0;
   }

   if (bytelen != NULL) {
      * bytelen = widelen * 2;
   }
   return bstr;
}

// Converts a BSTR to plain char* in UTF-8.
static char * BSTR_to_char (BSTR bstr) {
   int    buflen;
   char * retvalue;
   int    ret;

   if (bstr != NULL) {
      // First find out the length we need for the return value.
      buflen = WideCharToMultiByte(CP_UTF8, 0, bstr, -1, NULL, 0, NULL, NULL);

      // Allocate buffer.
      New(902, retvalue, buflen + 1, char);

      // Get the goods
      ret = WideCharToMultiByte(CP_UTF8, 0, bstr, -1, retvalue, buflen, NULL, NULL);

      if (! ret) {
         int err = GetLastError();
         croak("Internal error: WideCharToMultiByte failed with %ld. Buflen was %d", err, buflen);
      }

      return retvalue;
   }
   else {
      return NULL;
   }
}

// And this one takes the BSTR all the way to an SV. If not submitted, the
// string is assumed to be NULL-terminated.
static SV * BSTR_to_SV (BSTR  bstr,
                        int   bstrlen = -1) {
   int    buflen;
   char * tmp;
   int    ret;
   SV   * sv;

   if (bstr != NULL) {
      if (bstrlen != 0) {
         // First find out the length we need for the return value.
         buflen = WideCharToMultiByte(CP_UTF8, 0, bstr, bstrlen, NULL, 0, NULL, NULL);

         // Allocate buffer.
         New(902, tmp, buflen, char);

         // Get the goods
         ret = WideCharToMultiByte(CP_UTF8, 0, bstr, bstrlen, tmp, buflen, NULL, NULL);

         if (! ret) {
            int err = GetLastError();
            croak("Internal error: WideCharToMultiByte failed with %ld. Buflen was %d", err, buflen);
         }

         // If bstrlen was -1, then bstr is null-terminated, and so is tmp,
         // and buflen 1 too long.
         if (bstrlen == -1) {
            buflen--;
         }

         sv = newSVpvn(tmp, buflen);
         SvUTF8_on(sv);
         Safefree(tmp);
      }
      else {
         sv = newSVpvn("", 0);
      }
   }
   else {
      sv = NULL;
   }

   return sv;
}

//======================================================================
// The big init block. What follows are routines when someone says
// C<use Win32::SqlServer> for the first time.
//======================================================================

//------------------------------------------------------------------------
// Routines to set up the static array gbl_init_props.
//-----------------------------------------------------------------------

// A helper routine to get default for APPNAME.
static BSTR get_scriptname () {
   // Get the name of the script, taken from Perl var $0. This is used as
   // the default application name in SQL Server.

   SV* sv;

   if (sv = perl_get_sv("0", FALSE))
   {
      // Get script name into a BSTR.
      BSTR tmp = SV_to_BSTR(sv);
      BSTR scriptname;
      WCHAR *p;

      // But this name is full path, and we want only the trailing bit.
      if (p = wcsrchr(tmp, '/'))
         ++p;
      else if (p = wcsrchr(tmp, '\\'))
         ++p;
      else if (p = wcsrchr(tmp, ':'))
          ++p;
      else
          p = tmp;

      scriptname = SysAllocString(p);
      SysFreeString(tmp);
      return scriptname;
   }
   else {
      return NULL;
   }
}

// And another one to get the default for WSID.
static BSTR get_hostname() {
   BSTR hostname = SysAllocStringLen(NULL, 31);
   memset(hostname, 0, 60);
   GetEnvironmentVariable(L"COMPUTERNAME", hostname, 30);
   return hostname;
}

// Add a property to the global array.
static void add_init_property (const char *  name,
                               init_propsets propset_enum,
                               DBPROPID      propid,
                               VARTYPE       datatype,
                               BOOL          default_empty,
                               const WCHAR * default_str,
                               int           default_int,
                               int          &ix)
{

   // Check that we are not exceeding the global array. Note that the last
   // slot must be left unusued, as this is used as a stop condition!
   if (ix >= MAX_INIT_PROPERTIES - 1) {
      croak("Internal error: size of array for init properties exceeded");
   }

   // Increment property set counter.
   init_propset_info[propset_enum].no_of_props++;

   strcpy(gbl_init_props[ix].name, name);
   gbl_init_props[ix].propset_enum = propset_enum;
   gbl_init_props[ix].property_id  = propid;
   gbl_init_props[ix].datatype     = datatype;
   VariantInit(&gbl_init_props[ix].default_value);

   if (! default_empty) {
      gbl_init_props[ix].default_value.vt = datatype;

      switch (datatype) {
         case VT_BOOL :
            gbl_init_props[ix].default_value.boolVal = default_int;
            break;

         case VT_I2 :
            gbl_init_props[ix].default_value.iVal = default_int;
            break;

         case VT_UI2 :
            gbl_init_props[ix].default_value.uiVal = default_int;
            break;

         case VT_I4 :
            gbl_init_props[ix].default_value.lVal = default_int;
            break;

         case VT_BSTR :
            gbl_init_props[ix].default_value.bstrVal = SysAllocString(default_str);
            break;

         default :
            croak ("Internal error: add_init_property was called witn unhandled vartype %d",
                    datatype);
            break;
       }
    }

    // And increase the index.
    ix++;
}

// And this is the routine that sets up the array.
static void setup_init_properties ()
{
   int ix = 0;
   BSTR scriptname = get_scriptname();
   BSTR hostname   = get_hostname();

   // Init array so that all entrys are unused and init propset_info.
   memset(gbl_init_props, not_in_use,
          MAX_INIT_PROPERTIES * sizeof(init_property));


   // DBPROPSET_DBINIT, main OLE DB init and auth properties.
   init_propset_info[oleinit_props].start = ix;
   init_propset_info[oleinit_props].no_of_props = 0;

   add_init_property("IntegratedSecurity", oleinit_props, DBPROP_AUTH_INTEGRATED,
                     VT_BSTR, FALSE, L"SSPI", NULL, ix);
   add_init_property("Password", oleinit_props, DBPROP_AUTH_PASSWORD,
                     VT_BSTR, TRUE, NULL, NULL, ix);
   add_init_property("Username", oleinit_props, DBPROP_AUTH_USERID,
                     VT_BSTR, TRUE, NULL, NULL, ix);
   add_init_property("Database", oleinit_props, DBPROP_INIT_CATALOG,
                     VT_BSTR, FALSE, L"tempdb", NULL, ix);
   add_init_property("Server", oleinit_props, DBPROP_INIT_DATASOURCE,
                     VT_BSTR, FALSE, L"(local)", NULL, ix);
   add_init_property("GeneralTimeout", oleinit_props, DBPROP_INIT_GENERALTIMEOUT,
                     VT_I4, FALSE, NULL, 0, ix);
   add_init_property("LCID", oleinit_props, DBPROP_INIT_LCID,
                      VT_I4, FALSE, NULL, GetUserDefaultLCID(), ix);
   add_init_property("Pooling", oleinit_props, DBPROP_INIT_OLEDBSERVICES,
                     VT_I4, FALSE, NULL, DBPROPVAL_OS_RESOURCEPOOLING, ix);
   add_init_property("Prompt", oleinit_props, DBPROP_INIT_PROMPT,
                     VT_I2, FALSE, NULL, DBPROMPT_NOPROMPT, ix);
   add_init_property("ConnectionString", oleinit_props, DBPROP_INIT_PROVIDERSTRING,
                     VT_BSTR, TRUE, NULL, NULL, ix);
   add_init_property("ConnectTimeout", oleinit_props, DBPROP_INIT_TIMEOUT,
                     VT_I4, FALSE, NULL, 15, ix);

   // DBPROPSET_SQLSERVERDBINIT, SQLOLEDB specific proprties.
   init_propset_info[ssinit_props].start = ix;
   init_propset_info[ssinit_props].no_of_props = 0;

   add_init_property("Appname", ssinit_props, SSPROP_INIT_APPNAME,
                     VT_BSTR, FALSE, scriptname, NULL, ix);
   add_init_property("Autotranslate", ssinit_props, SSPROP_INIT_AUTOTRANSLATE,
                     VT_BOOL, TRUE, NULL, NULL, ix);
   add_init_property("Language", ssinit_props, SSPROP_INIT_CURRENTLANGUAGE,
                     VT_BSTR, TRUE, NULL, NULL, ix);
   add_init_property("AttachFilename", ssinit_props, SSPROP_INIT_FILENAME,
                     VT_BSTR, TRUE, NULL, NULL, ix);
   add_init_property("NetworkAddress", ssinit_props, SSPROP_INIT_NETWORKADDRESS,
                     VT_BSTR, TRUE, NULL, NULL, ix);
   add_init_property("Netlib", ssinit_props, SSPROP_INIT_NETWORKLIBRARY,
                     VT_BSTR, TRUE, NULL, NULL, ix);
   add_init_property("PacketSize", ssinit_props, SSPROP_INIT_PACKETSIZE,
                     VT_I4, TRUE, NULL, NULL, ix);
   add_init_property("UseProcForPrep", ssinit_props, SSPROP_INIT_USEPROCFORPREP,
                     VT_I4, FALSE, NULL, SSPROPVAL_USEPROCFORPREP_OFF, ix);
   add_init_property("Hostname", ssinit_props, SSPROP_INIT_WSID,
                     VT_BSTR, FALSE, hostname, NULL, ix);
   // Available first in 2.6.
   add_init_property("Encrypt", ssinit_props, SSPROP_INIT_ENCRYPT,
                     VT_BOOL, TRUE, NULL, NULL, ix);

   // The above properties are those that are in SQLOLEDB.
   no_of_sqloledb_ssprops = init_propset_info[ssinit_props].no_of_props;

   // These properties must come last, because there are supported by SQLNCLI
   // only.
   add_init_property("FailoverPartner", ssinit_props, SSPROP_INIT_FAILOVERPARTNER,
                     VT_BSTR, TRUE, NULL, NULL, ix);
   add_init_property("TrustServerCert", ssinit_props, SSPROP_INIT_TRUST_SERVER_CERTIFICATE,
                     VT_BOOL, TRUE, NULL, NULL, ix);
   add_init_property("OldPassword", ssinit_props, SSPROP_AUTH_OLD_PASSWORD,
                     VT_BSTR, TRUE, NULL, NULL, ix);

   // DBPROPSET_DATASOURCE, data-source properties.
   init_propset_info[datasrc_props].start = ix;
   init_propset_info[datasrc_props].no_of_props = 0;

   add_init_property("MultiConnections", datasrc_props, DBPROP_MULTIPLECONNECTIONS,
                     VT_BOOL, FALSE, NULL, FALSE, ix);

   SysFreeString(scriptname);
   SysFreeString(hostname);
}

//---------------------------------------------------------------------
// Routines for setting up and reading the type map. The first two are
// called on initialization only.
//---------------------------------------------------------------------
static void add_type_entry(const char * name,
                           DBTYPE       indicator,
                           int         &ix)
{
   int new_ix = ix + strlen(name) + 1;

   if (ix + strlen(name) + 1 > SIZE_TYPE_MAP) {
       croak ("Internal error: Adding %s at index ix = %d exceeds the size %d of the type map",
               name, ix, SIZE_TYPE_MAP);
   }
   strcpy(&(typenames[ix]), name);
   typenames[new_ix - 1] = ' ';
   typeindicators[ix] = indicator;
   ix = new_ix;
}

static void fill_type_map ()
{
   int    ix  = 1;

   typenames[0] = ' ';
   memset(typeindicators, 0, sizeof(DBTYPE) * SIZE_TYPE_MAP);
   add_type_entry("bigint",           DBTYPE_I8, ix);
   add_type_entry("binary",           DBTYPE_BYTES, ix);
   add_type_entry("bit",              DBTYPE_BOOL, ix);
   add_type_entry("char",             DBTYPE_STR, ix);
   add_type_entry("datetime",         DBTYPE_DBTIMESTAMP, ix);
   add_type_entry("decimal",          DBTYPE_NUMERIC, ix);
   add_type_entry("float",            DBTYPE_R8, ix);
   add_type_entry("image",            DBTYPE_BYTES, ix);
   add_type_entry("int",              DBTYPE_I4, ix);
   add_type_entry("money",            DBTYPE_CY, ix);
   add_type_entry("nchar",            DBTYPE_WSTR, ix);
   add_type_entry("ntext",            DBTYPE_WSTR, ix);
   add_type_entry("numeric",          DBTYPE_NUMERIC, ix);
   add_type_entry("nvarchar",         DBTYPE_WSTR, ix);
   add_type_entry("real",             DBTYPE_R4, ix);
   add_type_entry("smalldatetime",    DBTYPE_DBTIMESTAMP, ix);
   add_type_entry("smallint",         DBTYPE_I2, ix);
   add_type_entry("smallmoney",       DBTYPE_CY, ix);
   add_type_entry("sql_variant",      DBTYPE_SQLVARIANT, ix);
   add_type_entry("text",             DBTYPE_STR, ix);
   add_type_entry("timestamp",        DBTYPE_BYTES, ix);
   add_type_entry("tinyint",          DBTYPE_UI1, ix);
   add_type_entry("uniqueidentifier", DBTYPE_GUID, ix);
   add_type_entry("UDT",              DBTYPE_UDT, ix);
   add_type_entry("varbinary",        DBTYPE_BYTES, ix);
   add_type_entry("varchar",          DBTYPE_STR, ix);
   add_type_entry("xml",              DBTYPE_XML, ix);

   typenames[ix] = '\0';
}

// And this routine looks up a name in the type map.
DBTYPE lookup_type_map(const char * nameoftype)
{
   char * tmp;
   int ix;

   New(902, tmp, strlen(nameoftype) + 10, char);

   sprintf(tmp, " %s ", nameoftype);
   char * hit = strstr(typenames, tmp);
   Safefree(tmp);

   if (hit == NULL) {
      return DBTYPE_EMPTY;
   }
   ix = hit + 1 - typenames;
   return typeindicators[ix];
}

//---------------------------------------------------------------------
// Initialization and finalization.
//--------------------------------------------------------------------

//-------------------------------------------------------------------
// Windows calls DllMain the DLL is (un)loaded. We need a critical
// section in initialize (which is called by Perl on use of the module),
// so that only the first process sets up the global structures.
//-------------------------------------------------------------------
static CRITICAL_SECTION CS;

BOOL WINAPI DllMain(
  HINSTANCE hinstDLL,     // handle to the DLL module
  DWORD    fdwReason,     // reason for calling function
  LPVOID   lpvReserved)   // reserved
{
  switch (fdwReason) {
     case DLL_PROCESS_ATTACH:
        InitializeCriticalSection(&CS);
        break;
     case DLL_PROCESS_DETACH:
        DeleteCriticalSection(&CS);
        break;
     default:
        break;
  }
  return TRUE;
}

// Called when a Perl script says C<use Win32::SqlServer>.
void initialize ()
{
   SV *sv;
   DWORD       err;
   HRESULT     ret = S_OK;
   char        obj[200];

   // In the critical section we create our starting point, the pointer to
   // OLE DB services. We also create a pointer to a conversion object.
   // Thess pointer will never be released.
   EnterCriticalSection(&CS);

   // Get classIDs for the two possible providers.
   if (IsEqualCLSID(clsid_sqloledb, CLSID_NULL) &&
       IsEqualCLSID(clsid_sqlncli, CLSID_NULL)) {

      ret = CLSIDFromProgID(L"SQLOLEDB", &clsid_sqloledb);
      if (FAILED(ret)) {
         clsid_sqloledb = CLSID_NULL;
      }

      ret = CLSIDFromProgID(L"SQLNCLI", &clsid_sqlncli);
      if (FAILED(ret)) {
         clsid_sqlncli = CLSID_NULL;
      }
   }

   if (OLE_malloc_ptr == NULL)
      CoGetMalloc(1, &OLE_malloc_ptr);

   if (data_init_ptr == NULL) {
      CoInitializeEx(NULL, COINIT_MULTITHREADED);

      ret = CoCreateInstance(CLSID_MSDAINITIALIZE, NULL, CLSCTX_INPROC_SERVER,
                             IID_IDataInitialize,
                             reinterpret_cast<LPVOID *>(&data_init_ptr));
      if (FAILED(ret)) {
         sprintf(obj, "IDataInitialize");
      }

      // Fill the type map and the default login properties here.
      fill_type_map();
      setup_init_properties();

#ifdef FILEDEBUG
      // Open debug file.
      if (dbgfile == NULL) {
         dbgfile = _wfopen(L"C:\\temp\\ut.txt", L"wbc");
         fprintf(dbgfile, "\xFF\xFE");
      }
#endif
   }
   if (SUCCEEDED(ret) && data_convert_ptr == NULL) {
      ret = CoCreateInstance(CLSID_OLEDB_CONVERSIONLIBRARY,
                             NULL, CLSCTX_INPROC_SERVER,
                             IID_IDataConvert,
                             (void **) &data_convert_ptr);
      if (FAILED(ret)) {
         sprintf(obj, "IDataConvert");
      }
   }

   LeaveCriticalSection(&CS);

   if (FAILED(ret)) {
      err = GetLastError();
      warn("Could not create '%s' object: %d", obj, err);
      warn("This could be because you don't have the MDAC on your machine,\n");
      warn("or an MDAC version you have is too arcane and not supported by\n");
      croak("Win32::SqlServer, which requires MDAC 2.6\n");
   }

   // Set Version string.
   if (sv = perl_get_sv("Win32::SqlServer::Version", TRUE))
   {
        char buff[256];
        sprintf(buff, "This is Win32::SqlServer, version %s\n\nCopyright (c) 2005 Erland Sommarskog\n",
                XS_VERSION);
        sv_setnv(sv, atof(XS_VERSION));
        sv_setpv(sv, buff);
        SvNOK_on(sv);
   }
}

//=======================================================================
// Diagnostics routines used at creation of an OlleDB object.
//=======================================================================

// Dumps the contents of a property array in case of an error
static void dump_properties(DBPROP init_properties[MAX_INIT_PROPERTIES],
                            BOOL   props_debug)
{
  BOOL too_old_sqloledb = FALSE;

  for (int i = 0; gbl_init_props[i].propset_enum != not_in_use; i++) {
       if (! props_debug &&
           init_properties[i].dwStatus == DBPROPSTATUS_OK)
           continue;

       char ststxt[50];
       switch (init_properties[i].dwStatus) {
          case DBPROPSTATUS_OK :
               sprintf(ststxt, "DBPROPSTATUS_OK"); break;
          case DBPROPSTATUS_BADCOLUMN :
               sprintf(ststxt, "DBPROPSTATUS_BADCOLUMN"); break;
          case DBPROPSTATUS_BADOPTION :
               sprintf(ststxt, "DBPROPSTATUS_BADOPTION"); break;
          case DBPROPSTATUS_BADVALUE :
               sprintf(ststxt, "DBPROPSTATUS_BADVALUE"); break;
          case DBPROPSTATUS_CONFLICTING :
               sprintf(ststxt, "DBPROPSTATUS_CONFLICTING"); break;
          case DBPROPSTATUS_NOTALLSETTABLE :
               sprintf(ststxt, "DBPROPSTATUS_NOTALLSETTABLE"); break;
          case DBPROPSTATUS_NOTAVAILABLE :
               sprintf(ststxt, "DBPROPSTATUS_NOTAVAILABLE"); break;
          case DBPROPSTATUS_NOTSET :
               sprintf(ststxt, "DBPROPSTATUS_NOTSET"); break;
          case DBPROPSTATUS_NOTSETTABLE :
               sprintf(ststxt, "DBPROPSTATUS_NOTSETTABLE"); break;
          case DBPROPSTATUS_NOTSUPPORTED :
               sprintf(ststxt, "DBPROPSTATUS_NOTSUPPORTED"); break;
          case -1 :
               sprintf(ststxt, "(not set by OLE DB provider)");
               too_old_sqloledb = TRUE;
               break;
       }
       PerlIO_printf(PerlIO_stderr(), "Property '%s', Status: %s, Value: ",
                      gbl_init_props[i].name, ststxt);
       if (init_properties[i].vValue.vt == VT_EMPTY) {
           PerlIO_printf(PerlIO_stderr(), "VT_EMPTY");
       }
       else {
          switch (gbl_init_props[i].datatype) {
             case VT_BOOL :
                PerlIO_printf(PerlIO_stderr(), "%d",
                              init_properties[i].vValue.boolVal);
                break;

             case VT_I2 :
                PerlIO_printf(PerlIO_stderr(), "%d",
                              init_properties[i].vValue.iVal);
                break;

             case VT_I4 :
                PerlIO_printf(PerlIO_stderr(), "%d",
                              init_properties[i].vValue.lVal);
                break;

             case VT_BSTR : {
                char * str = BSTR_to_char(init_properties[i].vValue.bstrVal);
                PerlIO_printf(PerlIO_stderr(), "'%s'", str);
                Safefree(str);
                break;
            }

            default :
                PerlIO_printf(PerlIO_stderr(), "UNKNOWN DATATYPE");
                break;
           }
       }

       PerlIO_printf(PerlIO_stderr(), ".\n");
   }

   if (too_old_sqloledb) {
      warn("The fact that status for one or more properties were not set by\n");
      warn("by the OLE DB provider, indicates that you are running an unsupported\n");
      warn("version of SQLOLEDB. To use Win32::SqlServer you must have at least\n");
      croak("version 2.6 of the MDAC, or you must use SQL Native Client\n");
   }
}

// This is purely a debug routine which is available, mainly to check for
// leaks, but is normally not called from anywhere.
void dump_internaldata(internaldata * mydata)
{
   dump_properties(mydata->init_properties, TRUE);

   warn("init_ptr = %x.\n", mydata->init_ptr);
   warn("datasrc_ptr = %x.\n", mydata->datasrc_ptr);
   warn("isautoconnected = %d.\n", mydata->isautoconnected);
   warn("provider = %d.\n", mydata->provider);
   warn("pending_cmd = %d.\n", mydata->pending_cmd);
   warn("paramfirst = %x.\n", mydata->paramfirst);
   warn("paramlast = %x.\n", mydata->paramlast);
   warn("no_of_params = %d.\n", mydata->no_of_params);
   warn("no_of_out_params = %d.\n", mydata->no_of_out_params);
   warn("params_available = %d.\n", mydata->params_available);
   warn("session_ptr = %x.\n", mydata->session_ptr);
   warn("cmdtext_ptr = %x.\n", mydata->cmdtext_ptr);
   warn("paramcmd_ptr = %x.\n", mydata->paramcmd_ptr);
   warn("ss_paramcmd_ptr = %x.\n", mydata->ss_paramcmd_ptr);
   warn("paramaccess_ptr = %x.\n", mydata->paramaccess_ptr);
   warn("all_params_OK = %d.\n", mydata->all_params_OK);
   warn("param_info = %x.\n", mydata->param_info);
   warn("param_bindings = %x.\n", mydata->param_bindings);
   warn("param_buffer = %x.\n", mydata->param_buffer);
   warn("size_param_buffer = %d.\n", mydata->size_param_buffer);
   warn("param_accessor = %d.\n", mydata->param_accessor);
   warn("param_bind_status = %x.\n", mydata->param_bind_status);
   warn("results_ptr = %x.\n", mydata->results_ptr);
   warn("have_resultset = %d.\n", mydata->have_resultset);
   warn("rowset_ptr = %x.\n", mydata->rowset_ptr);
   warn("rowaccess_ptr = %x.\n", mydata->rowaccess_ptr);
   warn("row_accessor = %d.\n", mydata->row_accessor);
   warn("column_keys = %x.\n", mydata->column_keys);
   warn("rowbuffer = %x.\n", mydata->rowbuffer);
   warn("rows_in_buffer = %d.\n", mydata->rows_in_buffer);
   warn("current_rowno = %d.\n", mydata->current_rowno);
   warn("no_of_cols = %d.\n", mydata->no_of_cols);
   warn("column_info = %x.\n", mydata->column_info);
   warn("colname_buffer = %x.\n", mydata->colname_buffer);
   warn("col_bindings = %x.\n", mydata->col_bindings);
   warn("col_bind_status = %x.\n", mydata->col_bind_status);
   warn("data_buffer = %x.\n", mydata->data_buffer);
   warn("size_data_buffer = %d.\n", mydata->size_data_buffer);
}

//=================================================================
// Creating and destroying OlleDB objects.
//=================================================================

// We release pointers a lot, so we have a macro that does it all.
#define free_ole_ptr(oleptr) \
   if (oleptr != NULL) { \
      oleptr->Release(); \
      oleptr = NULL; \
   } \


// This routine allocates an internaldata structure and returns the pointer
// as an integer value.
static int setupinternaldata()
{
    internaldata  * mydata;  // Pointer to area for internal data.

    // Create struct for pointers we need to keep between calls, and initiate
    // all pointers to NULL.
    New(902, mydata, 1, internaldata);
    mydata->isautoconnected   = FALSE;
    mydata->provider          = (IsEqualCLSID(clsid_sqlncli, CLSID_NULL) ?
                                 provider_sqloledb : provider_sqlncli);
    mydata->init_ptr          = NULL;
    mydata->pending_cmd       = NULL;
    mydata->paramfirst        = NULL;
    mydata->paramlast         = NULL;
    mydata->no_of_params      = 0;
    mydata->no_of_out_params  = 0;
    mydata->params_available  = FALSE;
    mydata->datasrc_ptr       = NULL;
    mydata->session_ptr       = NULL;
    mydata->cmdtext_ptr       = NULL;
    mydata->paramcmd_ptr      = NULL;
    mydata->ss_paramcmd_ptr   = NULL;
    mydata->paramaccess_ptr   = NULL;
    mydata->all_params_OK     = TRUE;
    mydata->param_info        = NULL;
    mydata->param_bindings    = NULL;
    mydata->param_buffer      = NULL;
    mydata->size_param_buffer = 0;
    mydata->param_accessor    = NULL;
    mydata->param_bind_status = NULL;
    mydata->results_ptr       = NULL;
    mydata->have_resultset    = FALSE;
    mydata->rowset_ptr        = NULL;
    mydata->rowaccess_ptr     = NULL;
    mydata->row_accessor      = NULL;
    mydata->column_keys       = NULL;
    mydata->rowbuffer         = NULL;
    mydata->rows_in_buffer    = 0;
    mydata->current_rowno     = 0;
    mydata->no_of_cols        = NULL;
    mydata->column_info       = NULL;
    mydata->colname_buffer    = NULL;
    mydata->col_bindings      = NULL;
    mydata->col_bind_status   = NULL;
    mydata->data_buffer       = NULL;
    mydata->size_data_buffer  = 0;


    // Set up the init property sets. First the GUIDs.
    mydata->init_propsets[oleinit_props].guidPropertySet =
            DBPROPSET_DBINIT;
    mydata->init_propsets[ssinit_props].guidPropertySet  =
            DBPROPSET_SQLSERVERDBINIT;
    mydata->init_propsets[datasrc_props].guidPropertySet =
            DBPROPSET_DATASOURCE;

    // Then number and pointer to the arrays.
    for (int i = 0; i <= NO_OF_INIT_PROPSETS; i++) {
       mydata->init_propsets[i].cProperties  = init_propset_info[i].no_of_props;
       mydata->init_propsets[i].rgProperties =
           &(mydata->init_properties[init_propset_info[i].start]);
    }

    // Then copy the properties from the global default properties.
    for (int j = 0; gbl_init_props[j].propset_enum != not_in_use; j++) {
       DBPROP  &prop = mydata->init_properties[j];
       prop.dwPropertyID = gbl_init_props[j].property_id;
       prop.dwOptions    = DBPROPOPTIONS_REQUIRED;
       prop.colid        = DB_NULLID;
       prop.dwStatus     = DBPROPSTATUS_OK;
       VariantInit(&prop.vValue);
       VariantCopy(&prop.vValue, &gbl_init_props[j].default_value);
    }

    return (int) mydata;
}

// This routine frees about allocation for receiving a result. It is
// called by nextrow, when there aer no more rows, or by cancelresultset.
void free_resultset_data(internaldata *mydata) {
   HRESULT ret;

   if (mydata->column_info != NULL) {
      OLE_malloc_ptr->Free(mydata->column_info);
      mydata->column_info = NULL;
   }
   if (mydata->colname_buffer != NULL) {
      OLE_malloc_ptr->Free(mydata->colname_buffer);
      mydata->colname_buffer = NULL;
   }
   if (mydata->col_bindings != NULL) {
      Safefree(mydata->col_bindings);
      mydata->col_bindings = NULL;
   }
   if (mydata->col_bind_status != NULL) {
      Safefree(mydata->col_bind_status);
      mydata->col_bind_status = NULL;
   }
   if (mydata->data_buffer != NULL) {
      Safefree(mydata->data_buffer);
      mydata->data_buffer = NULL;
   }

   if (mydata->rowbuffer != NULL) {
      ret = mydata->rowset_ptr->ReleaseRows(mydata->rows_in_buffer,
                                            mydata->rowbuffer,
                                            NULL, NULL, NULL);
      if (FAILED(ret)) {
         croak("rowset_ptr->ReleaseRows failed with %08X.\n", ret);
      }
      Safefree(mydata->rowbuffer);
      mydata->rowbuffer = NULL;
   }
   mydata->rows_in_buffer = 0;
   mydata->current_rowno = 0;

   if (mydata->row_accessor != NULL) {
      if (mydata->rowaccess_ptr != NULL) {
         mydata->rowaccess_ptr->ReleaseAccessor(mydata->row_accessor, NULL);
      }
      mydata->row_accessor = NULL;
   }

   if (mydata->column_keys != NULL) {
      for (ULONG i = 0; i < mydata->no_of_cols; i++) {
         if (mydata->column_keys[i] && SvOK(mydata->column_keys[i])) {
            SvREFCNT_dec(mydata->column_keys[i]);
         }
      }
      Safefree(mydata->column_keys);
      mydata->column_keys = NULL;
   }
   mydata->no_of_cols = 0;

   free_ole_ptr(mydata->rowaccess_ptr);
   free_ole_ptr(mydata->rowset_ptr);
   mydata->have_resultset = FALSE;
}

// This routine frees up everything with a saved parameter list. Normally
// called before we execute a parameterised command. Also called from
// free_batch_data as a safety precaution.
void free_pending_cmd(internaldata *mydata) {
   if (mydata->pending_cmd != NULL) {
      SysFreeString(mydata->pending_cmd);
      mydata->pending_cmd = NULL;
   }

   while (mydata->paramfirst != NULL) {
      paramdata * tmp;
      BYTE      * provider_ptr = NULL;
      tmp = mydata->paramfirst;

      // If there is a parameter buffer, there might be a pointer to an area
      // allocated by the provider.
      if (mydata->param_buffer != NULL) {
         provider_ptr = *(BYTE **) &mydata->param_buffer[tmp->binding.obValue];
      }

      SysFreeString(tmp->param_info.pwszName);
      SysFreeString(tmp->param_info.pwszDataSourceType);

      // buffer_ptr is a saved address to input parameter.
      if (tmp->buffer_ptr != NULL) {
         // So the area in the parameter buffer is a pointer, and if its not
         // NULL or the same as our save pointer, we must free the area.
         if (provider_ptr != NULL && provider_ptr != tmp->buffer_ptr) {
            OLE_malloc_ptr->Free(provider_ptr);
         }
         Safefree(tmp->buffer_ptr);
      }

      // bstr is a saved addres to an nvarchar parameter, different pool than
      // buffer_ptr.
      if (tmp->bstr != NULL) {
         if (provider_ptr != NULL && provider_ptr != (BYTE *) tmp->bstr) {
            OLE_malloc_ptr->Free(provider_ptr);
         }
         SysFreeString(tmp->bstr);
      }

      // Any parameter properties must be released.
      if (tmp->param_props_cnt > 0) {
         for (int ix = 0; ix < tmp->param_props_cnt; ix++) {
            VariantClear(&tmp->param_props[ix].vValue);
         }
         Safefree(tmp->param_props);
      }

      mydata->paramfirst = tmp->next;
      Safefree(tmp);
   }

   mydata->paramlast = NULL;
}

// This routine is called whenever we need to cancel everything allocated
// for a query batch.
void free_batch_data(internaldata *mydata) {
   // First free eveything associated with a result.
   free_resultset_data(mydata);
   mydata->params_available = FALSE;

   // Pending command.
   free_pending_cmd(mydata);

   // Parameter information.
   if (mydata->param_info != NULL) {
      Safefree(mydata->param_info);
      mydata->param_info = NULL;
   }

   if (mydata->param_bindings != NULL) {
      Safefree(mydata->param_bindings);
      mydata->param_bindings = NULL;
   }

   if (mydata->param_buffer != NULL) {
      Safefree(mydata->param_buffer);
      mydata->param_buffer = NULL;
   }

   mydata->size_param_buffer = 0;
   mydata->no_of_params = 0;
   mydata->no_of_out_params = 0;
   mydata->all_params_OK = TRUE;

   if (mydata->param_accessor != NULL) {
      if (mydata->paramaccess_ptr != NULL) {
         HRESULT ret;
         ret = mydata->paramaccess_ptr->ReleaseAccessor(
                                        mydata->param_accessor, NULL);
         if (FAILED(ret)) {
            croak("paramaccess_ptr->ReleaseAccessor failed with %08X.\n", ret);
         }
      }
      mydata->param_accessor = NULL;
   }

   if (mydata->param_bind_status != NULL) {
      Safefree(mydata->param_bind_status);
      mydata->param_bind_status = NULL;
   }

   free_ole_ptr(mydata->results_ptr);
   free_ole_ptr(mydata->paramaccess_ptr);
   free_ole_ptr(mydata->paramcmd_ptr);
   free_ole_ptr(mydata->ss_paramcmd_ptr);
   free_ole_ptr(mydata->cmdtext_ptr);
   free_ole_ptr(mydata->session_ptr);

   // Release the data-source pointer only if auto-connected.
   if (mydata->isautoconnected) {
      free_ole_ptr(mydata->datasrc_ptr);
      free_ole_ptr(mydata->init_ptr);
   }
}


//=====================================================================
// Routines for checking and reporting errors.
//=====================================================================
// olle_croak calls croak, but before that it calls free_batch_data to
// release all that is allocated.
static void olle_croak(SV         * olle_ptr,
                       const char * msg,
                       ...)
{
    va_list args;

    if (olle_ptr != NULL) {
       free_batch_data(get_internaldata(olle_ptr));
    }

    va_start(args, msg);
    vcroak(msg, &args);
    va_end(args);     // Not reached.
}

// msg_handler invokes the user-defined callback, or the default built-in one.
// Most errors comes from SQL Server, which is reflected in the interface,
// but Win32::SqlServer can use it for its own errors too.
static void msg_handler (SV        *olle_ptr,
                         int        msgno,
                         int        msgstate,
                         int        severity,
                         BSTR       msgtext,
                         LPOLESTR   srvname,
                         LPOLESTR   procname,
                         ULONG      line,
                         LPOLESTR   sqlstate,
                         LPOLESTR   source,
                         ULONG      n,
                         ULONG      no_of_errs)
{
    SV ** callback_ptr;
    SV *  callback = NULL;

    // Note that if error occurs during login, we may not yet have an
    // olle_ptr;

    if (olle_ptr && SvOK(olle_ptr)) {
       if (callback_ptr = fetch_from_hash(olle_ptr, HV_msgcallback)) {
          callback = * callback_ptr;
          SvGETMAGIC(callback);
       }
    }

    if (callback && SvOK(callback))  {  // a perl error handler has been installed */
        dSP;
        int  retval;
        int  count;
        SV * sv_srvname;
        SV * sv_msgtext;
        SV * sv_procname;
        SV * sv_sqlstate;
        SV * sv_source;

        PUSHMARK(sp);
        ENTER;
        SAVETMPS;

        // Push a copy of the Perl ptr to the stack.
        XPUSHs(sv_mortalcopy(olle_ptr));

        XPUSHs(sv_2mortal (newSViv (msgno)));
        XPUSHs(sv_2mortal (newSViv (msgstate)));
        XPUSHs(sv_2mortal (newSViv (severity)));

        if (SysStringLen(msgtext) > 0) {
            sv_msgtext = BSTR_to_SV(msgtext);
            XPUSHs(sv_2mortal(sv_msgtext));
        }
        else
            XPUSHs(&PL_sv_undef);

        if (srvname && wcslen(srvname) > 0) {
            sv_srvname = BSTR_to_SV(srvname);
            XPUSHs(sv_2mortal(sv_srvname));
        }
        else
            XPUSHs(&PL_sv_undef);

        if (procname && wcslen(procname) > 0) {
           sv_procname = BSTR_to_SV(procname);
           XPUSHs(sv_2mortal (sv_procname));
        }
        else
            XPUSHs(&PL_sv_undef);

        XPUSHs(sv_2mortal (newSViv (line)));

        if (sqlstate && wcslen(sqlstate) > 0) {
           sv_sqlstate = BSTR_to_SV(sqlstate);
           XPUSHs(sv_2mortal(sv_sqlstate));
        }
        else
           XPUSHs(&PL_sv_undef);

        if (source && wcslen(source) > 0) {
           sv_source = BSTR_to_SV(source);
           XPUSHs(sv_2mortal(sv_source));
        }
        else
           XPUSHs(&PL_sv_undef);

        XPUSHs(sv_2mortal (newSViv (n)));
        XPUSHs(sv_2mortal (newSViv (no_of_errs)));

        PUTBACK;
        if ((count = call_sv(callback, G_SCALAR)) != 1)
            croak("A msg handler cannot return a LIST");
        SPAGAIN;
        retval = POPi;

        PUTBACK;
        FREETMPS;
        LEAVE;

        if (retval == 0) {
           olle_croak(olle_ptr, "Terminating on fatal error");
        }
    }
    else {
       // Here follows the XS message handler.

       // Only print complete infomation for errors.
       if (severity >= 11)  {
          if (source && wcslen(source) > 0) {
             char * charstr = BSTR_to_char(source);
             if (strlen(charstr) > 0)
                PerlIO_printf(PerlIO_stderr(), "Source %s\n", charstr);
             Safefree(charstr);
          }
          if (srvname && wcslen(srvname) > 0) {
             char * charstr = BSTR_to_char(srvname);
             if (strlen(charstr) > 0)
                PerlIO_printf(PerlIO_stderr(), "Server %s, ", charstr);
             Safefree(charstr);
          }
          PerlIO_printf(PerlIO_stderr(),"Msg %ld, Level %d, State %d",
                        msgno, severity, msgstate);
          if (procname && wcslen(procname) > 0) {
             char * charstr = BSTR_to_char(procname);
             if (strlen(charstr) > 0)
                PerlIO_printf(PerlIO_stderr(), ", Procedure '%s'", charstr);
             Safefree(charstr);
          }
          if (line > 0)
              PerlIO_printf(PerlIO_stderr(), ", Line %d", line);
          PerlIO_printf(PerlIO_stderr(), "\n\t");
       }

       if (SysStringLen(msgtext) > 0) {
           char *  charstr = BSTR_to_char(msgtext);
           PerlIO_printf(PerlIO_stderr(), "%s\n", charstr);
           Safefree(charstr);
       }
       else {
           PerlIO_printf(PerlIO_stderr(), "\n");
       }
    }
}

// A wrapper on msg_handler to produce a message with OlleDB as source.
static void olledb_message (SV    * olle_ptr,
                            int     msgno,
                            int     state,
                            int     severity,
                            BSTR    msg)
{
   msg_handler(olle_ptr, msgno, state, severity, msg,
               NULL, NULL, 0, NULL, L"Win32::SqlServer", 1, 1);
}

// The same with msg in 8-bit, this one is called from Perl code.
static void olledb_message (SV          * olle_ptr,
                            int           msgno,
                            int           state,
                            int           severity,
                            const char  * msg)
{
   BSTR  bstr_msg = SysAllocStringLen(NULL, strlen(msg) + 1);
   wsprintf(bstr_msg, L"%S", msg);
   olledb_message(olle_ptr, msgno, state, severity, bstr_msg);
   SysFreeString(bstr_msg);
}



// Dumps the contents of error_info_obj. This is a helper to check_for_errors
// below.
static void dump_error_info(SV            * olle_ptr,
                            const char    * context,
                            const HRESULT   hresult,
                            BOOL            call_msg_handler,
                            IErrorInfo    * error_info_obj,
                            ERRORINFO     * error_info_rec)
{

   if (error_info_obj != NULL) {
      BSTR bstr_source;
      BSTR bstr_description;

      error_info_obj->GetSource(&bstr_source);
      error_info_obj->GetDescription(&bstr_description);

      if (call_msg_handler) {
         BSTR bstr_context = SysAllocStringLen(NULL, strlen(context) + 1);
         BSTR hres_str = SysAllocStringLen(NULL, 11);
         wsprintf(bstr_context, L"%S", context);
         wsprintf(hres_str, L"%08x", hresult);

         msg_handler(olle_ptr, -1, 127, 16, bstr_description, NULL,
                     bstr_context, 0,
                     hres_str, bstr_source, 1, 1);

         SysFreeString(hres_str);
         SysFreeString(bstr_context);
      }
      else {
         char * source = BSTR_to_char(bstr_source);
         char * description = BSTR_to_char(bstr_description);

         warn("Source '%s' said '%s'.\n", source, description);

         Safefree(source);
         Safefree(description);
      }

      SysFreeString(bstr_source);
      SysFreeString(bstr_description);
   }

   if (error_info_rec != NULL && ! call_msg_handler) {
      LPOLESTR uni_clsid_str;
      LPOLESTR uni_iid_str;
      char    *clsid_str;
      char    *iid_str;

      // To display the GUIDs, we first we need to format them as strings,
      // but as we get UTF-16 strings we then convert to UTF-8.
      StringFromCLSID(error_info_rec->clsid, &uni_clsid_str);
      StringFromIID(error_info_rec->iid, &uni_iid_str);
      clsid_str = BSTR_to_char(uni_clsid_str);
      iid_str = BSTR_to_char(uni_iid_str);
      warn("HRESULT: %08x, Minor: %d, CLSID: %s, Interface ID: %s, DispID %ld.\n",
            error_info_rec->hrError, error_info_rec->dwMinor,
            clsid_str, iid_str, error_info_rec->dispid);
      CoTaskMemFree(uni_clsid_str);
      CoTaskMemFree(uni_iid_str);
      Safefree(clsid_str);
      Safefree(iid_str);
   }
}

// This routine checks for errors. If there are errors that comes from SQL
// Server, we call msg_handle for each message, and msg_handle may call a
// customed-installed error handler. If any erros comes from SQLOLEDB,
// we consider them programming errors, and croak. The parameters context
// and hresult are included in the croak message. The routine only checks
// whether hersult is an error in the case that GetErrorInfo does not return
// anything.
static void check_for_errors(SV *          olle_ptr,
                             const char   *context,
                             const HRESULT hresult,
                             BOOL          dieonnosql)
{
   IErrorInfo*     error_info_main      = NULL;
   int             no_of_sqlerrs        = 0;
   IErrorRecords*  error_records        = NULL;
   ULONG           no_of_errs;
   LCID            our_locale = GetUserDefaultLCID();

   // The OLE DB documentation says that we should check the interface for
   // ISupportErrorInfo, before we call GetErrorInfo. However, it appears
   // that IMultipleResults does not support ISupportErrorInfo, and it is
   // here we get the SQL errors. So we approach GetErrorInfo directly.

   GetErrorInfo(0, &error_info_main);

   // Check first if got any error information.
   if (error_info_main == NULL) {
     if (FAILED(hresult)) {
        // There was no error message, but obviously things went wrong anyway.
        // There is no reason to carry on.
        olle_croak(olle_ptr,
           "Internal error: %s failed with %08x. No further error infoformation was collected",
           context, hresult);
     }
     // It seems that everything went just fine.
     return;
   }

   // If we come here we have an error_info_main. Try to get the detail records.
   error_info_main->QueryInterface(IID_IErrorRecords,
                                   (void **) &error_records);
   if (error_records == NULL) {
   // We did not, but we some error have occurred, and we are going do die.
   // And here we don't care about dieonnosql.
      dump_error_info(olle_ptr, context, hresult, FALSE, error_info_main, NULL);
      error_info_main->Release();
      olle_croak(olle_ptr, "Internal error: %s failed with %08x", context, hresult);
   }

   // Get number of errors.
   error_records->GetRecordCount(&no_of_errs);

   // Then loop over the errors backwards. That will gives at least the
   // SQL errors in the order SQL Server produces them.
   for (ULONG n = 0; n < no_of_errs; n++) {
      ULONG errorno = no_of_errs - n - 1;
      ISQLErrorInfo*  sql_error_info  = NULL;

      // Try to get customer objects, for SQL errors and SQL Server errors.
      error_records->GetCustomErrorObject(errorno, IID_ISQLErrorInfo,
                                         (IUnknown **) &sql_error_info);

      if (sql_error_info != NULL) {
      // This is an SQL error, so we will call the message handler, and
      // we will survive the ordeal.
         LONG                  msgno;
         BSTR                  sqlstate;
         SSERRORINFO         * sserror_rec = NULL;
         OLECHAR             * sserror_strings = NULL;
         ISQLServerErrorInfo*  sqlserver_error_info = NULL;

         no_of_sqlerrs++;

         // Get SQLstate and message number and convert to char *.
         sql_error_info->GetSQLInfo(&sqlstate, &msgno);

         // Now, get the SQL Server errors.
         sql_error_info->QueryInterface(IID_ISQLServerErrorInfo,
                                       (void **) &sqlserver_error_info);

         // We're done with this object.
         sql_error_info->Release();

         // See if there is any SQL Server information. Normally there is,
         // but not if there is an SQL error detected by SQLOLEDB.
         if (sqlserver_error_info != NULL) {
            sqlserver_error_info->GetErrorInfo(&sserror_rec,
                                               &sserror_strings);
            sqlserver_error_info->Release();
         }

         if (sserror_rec != NULL) {
         // This is a regular SQL error. Call msg_handler.
            msg_handler(olle_ptr, sserror_rec->lNative,
                        sserror_rec->bState, sserror_rec->bClass,
                        sserror_rec->pwszMessage, sserror_rec->pwszServer,
                        sserror_rec->pwszProcedure, sserror_rec->wLineNumber,
                        sqlstate, NULL, n + 1, no_of_errs);

            // Clean-up time.
            OLE_malloc_ptr->Free(sserror_rec);
            OLE_malloc_ptr->Free(sserror_strings);
         }
         else {
         // An SQL error detected by SQLOLEBB.
            BSTR         source;
            BSTR         description;
            IErrorInfo*  error_info_detail    = NULL;

            error_records->GetErrorInfo(errorno, our_locale, &error_info_detail);

            if (error_info_detail != NULL) {
               error_info_detail->GetSource(&source);
               error_info_detail->GetDescription(&description);
               // Call msg_handler, providing values for missing items.
               msg_handler(olle_ptr, msgno, 1, 16, description, NULL, NULL, 0,
                           sqlstate, source, n + 1, no_of_errs);

               SysFreeString(source);
               SysFreeString(description);
               error_info_detail->Release();
            }
            else {
            // Eh? Missing locale? Whatever, don't drop it on the floor.
               BSTR msg = SysAllocString(L"Error message missing");
               BSTR source = SysAllocString(L"Win32::SqlServer");
               msg_handler(olle_ptr, msgno, 0, 16, msg, NULL, NULL, 0,
                           sqlstate, source, n + 1, no_of_errs);
               SysFreeString(msg);
               SysFreeString(source);
            }

         }
         SysFreeString(sqlstate);
      }
      else {
      // This is could be an internal error. Then again, sometimes SQLOLEDB
      // does not do better than this.
         IErrorInfo*  error_info_detail   = NULL;
         ERRORINFO    error_rec;
         error_records->GetErrorInfo(errorno, our_locale, &error_info_detail);
         error_records->GetBasicErrorInfo(errorno, &error_rec);
         dump_error_info(olle_ptr, context, hresult, ! dieonnosql,
                         error_info_detail, &error_rec);
         error_info_detail->Release();
      }
   }

   error_records->Release();
   error_info_main->Release();

   if (no_of_sqlerrs == 0 && dieonnosql) {
   // Game over
      olle_croak(olle_ptr, "Internal error: %s failed with %08X",
                 context, hresult);
   }
}

// The default version of check_for_errors that dies on all non-SQL errors.
static void check_for_errors(SV *          olle_ptr,
                             const char   *context,
                             const HRESULT hresult)
{
    check_for_errors(olle_ptr, context, hresult, TRUE);
}


// This routine checks specifically for conversion errors. We look at
// return code and at DBSTATUS, but we don't try any error object. God
// knows whether IDataConvert supports that.
static void check_convert_errors (char*        msg,
                                  DBSTATUS     dbstatus,
                                  DBBINDSTATUS bind_status,
                                  HRESULT      ret)
{

   char    bad_status[200];
   BOOL    has_failed = TRUE;   // We clear in case we see a good status.

   switch (dbstatus) {
      case DBSTATUS_S_TRUNCATED :    // This merits only a warning.
         warn("Truncatation occured with '%s'", msg);
         // Fall-through
      case DBSTATUS_S_OK :
         has_failed  = FALSE;
         break;

      case DBSTATUS_S_ISNULL :
      // This should have been handled elsewhere, so if it comes where its
      // an error.
         sprintf(bad_status, "DBSTATUS_S_ISNULL");
         break;

      case DBSTATUS_E_BADACCESSOR :
          switch (bind_status) {
           case DBBINDSTATUS_OK :
              sprintf(bad_status, "DBSTATUS_E_BADACCESSOR/DBBINDSTATUS_OK");
              break;
           case DBBINDSTATUS_BADORDINAL :
              sprintf(bad_status, "DBSTATUS_E_BADACCESSOR/DBBINDSTATUS_BADORDINAL");
              break;
           case DBBINDSTATUS_UNSUPPORTEDCONVERSION :
              sprintf(bad_status, "DBSTATUS_E_BADACCESSOR/DBBINDSTATUS_UNSUPPORTEDCONVERSION");
              break;
           case DBBINDSTATUS_BADBINDINFO :
              sprintf(bad_status, "DBSTATUS_E_BADACCESSOR/DBBINDSTATUS_BADBINDINFO");
              break;
           case DBBINDSTATUS_BADSTORAGEFLAGS :
              sprintf(bad_status, "DBSTATUS_E_BADACCESSOR/DBBINDSTATUS_BADSTORAGEFLAGS");
              break;
           case DBBINDSTATUS_NOINTERFACE :
              sprintf(bad_status, "DBSTATUS_E_BADACCESSOR/DBBINDSTATUS_NOINTERFACE");
              break;
           default :
              sprintf(bad_status,
                      "DBSTATUS_E_BADACCESSOR/unidentified status %d", bind_status);
              break;
        }
        break;

      case DBSTATUS_E_CANTCONVERTVALUE :
          sprintf(bad_status, "DBSTATUS_E_CANTCONVERTVALUE");
          break;

      case DBSTATUS_E_CANTCREATE :
          sprintf(bad_status, "DBSTATUS_E_CANTCREATE");
          break;

      case DBSTATUS_E_DATAOVERFLOW :
          sprintf(bad_status, "DBSTATUS_E_DATAOVERFLOW");
          break;

      case DBSTATUS_E_SIGNMISMATCH :
          sprintf(bad_status, "DBSTATUS_E_SIGNMISMATCH");
          break;

      case DBSTATUS_E_UNAVAILABLE :
          sprintf(bad_status, "DBSTATUS_E_UNAVAILABLE");
          break;

      default :
          sprintf(bad_status, "Unidentified status value: %d", dbstatus);
          break;
    }

    if (has_failed) {
       if (FAILED(ret))
          warn("Operation '%s' failed with return status %d", msg, ret);
       croak("Operation '%s' gave bad status '%s'", msg, bad_status);
    }

    if (FAILED(ret)) {
       croak("Operation '%s' failed with return status %d", msg, ret);
    }
}

// Overloaded version with out bind_status.
static void check_convert_errors (char*        msg,
                                  DBSTATUS     dbstatus,
                                  HRESULT      ret)
{
    check_convert_errors(msg, dbstatus, DBBINDSTATUS_OK, ret);
}

//==================================================================
// Connection and disconnection, including setting properties.
//==================================================================

// Connect, called from $X->Connect() and $X->executebatch for autoconnect.
BOOL do_connect (SV    * olle_ptr,
                 BOOL    isautoconnect)
{
    internaldata  * mydata = get_internaldata(olle_ptr);
    HRESULT         ret      = S_OK;
    IDBProperties * property_ptr;
    CLSID         * clsid;
    char            provider_name[80];

    switch (mydata->provider) {
       // At this point provider_default should never appear.
       case provider_sqlncli :
         clsid = &clsid_sqlncli;
         sprintf(provider_name, "SQLNCLI");
         break;

       case provider_sqloledb :
         clsid = &clsid_sqloledb;
         sprintf(provider_name, "SQLOLEDB");
         break;

       default :
          croak ("Internal error: Illegal value %d for the provider enum",
                 mydata->provider);
    }
    if (FAILED(ret)) {
       croak("Chosen provider '%s' does not appear to be installed on this machine",
              provider_name);
    }

    ret = data_init_ptr->CreateDBInstance(*clsid,
                         NULL, CLSCTX_INPROC_SERVER,
                         NULL, IID_IDBInitialize,
                         reinterpret_cast<IUnknown **> (&mydata->init_ptr));
    if (FAILED(ret)) {
       croak("Internal error: IDataInitliaze->CreateDBInstance failed: %08X", ret);
    }

    // We need a property object.
    ret = mydata->init_ptr->QueryInterface(IID_IDBProperties,
                                           (void **) &property_ptr);
    if (FAILED(ret)) {
       croak("Internal error: init_ptr->QueryInterface to create Property object failed with hresult %x", ret);
    }

    // If we are using SQLOLEDB, we should reduce the number of SSPROPS,
    // because some are in Native Client only. Since there are old SQLOLEDB
    // we don't support, we have a special check for these.
    if (mydata->provider == provider_sqloledb) {
       mydata->init_propsets[ssinit_props].cProperties = no_of_sqloledb_ssprops;

       // Set all dwStatus to -1 for the first two propsets, this helps to
       // detect that some properties were not set, because we're in for an
       // old version of SQLOLEDB.
       for (int p = oleinit_props; p <= ssinit_props; p++) {
          for (UINT i = init_propset_info[p].start;
               i < mydata->init_propsets[p].cProperties +
                  init_propset_info[p].start; i++) {
             mydata->init_properties[i].dwStatus = -1;
          }
       }
    }

    ret = property_ptr->SetProperties(2, mydata->init_propsets);
    if (FAILED(ret)) {
       dump_properties(mydata->init_properties, OptPropsDebug(olle_ptr));
       croak("Internal error: property_ptr->SetProperties for initialization props failed with hresult %x", ret);
    }

    // This is the place where we actually log in to SQL Server. We might
    // be reusing a connection from a pool.
    ret = mydata->init_ptr->Initialize();

    // If success, continue with creating data-source object.
    if (SUCCEEDED(ret)) {
       // Set properties for the data source.
       ret = property_ptr->SetProperties(1, &mydata->init_propsets[datasrc_props]);
       check_for_errors(NULL, "property_ptr->SetProperties for data-source props",
                        ret);


       // Get a data source object.
       ret = mydata->init_ptr->QueryInterface(IID_IDBCreateSession,
                                            (void **) &(mydata->datasrc_ptr));
       check_for_errors(olle_ptr, "init_ptr->QueryInterface for data source",
                        ret);
       mydata->isautoconnected = isautoconnect;
    }
    else {
       dump_properties(mydata->init_properties, OptPropsDebug(olle_ptr));
       check_for_errors(olle_ptr, "init_ptr->Initialize", ret);
    }

    // And release the pointers.
    //init_ptr->Release();
    property_ptr->Release();

    return SUCCEEDED(ret);
}

// This is $X->setloginproperty.
void setloginproperty(SV   * olle_ptr,
                      char * prop_name,
                      SV   * prop_value)
{
   internaldata * mydata = get_internaldata(olle_ptr);
   int            ix = 0;

   // If we are connected, and warnings are enabled, emit a warning.
   if (mydata->datasrc_ptr != NULL) {
      olle_croak(olle_ptr, "You cannot set login properties while connected");
   }

   // Check we got a proper prop_name.
   if (prop_name == NULL) {
      croak("Property name must not be NULL.");
   }

   // Look up property name in the global array.
   while (gbl_init_props[ix].propset_enum != not_in_use &&
          _stricmp(prop_name, gbl_init_props[ix].name) != 0) {
      ix++;
   }

   if (gbl_init_props[ix].propset_enum == not_in_use) {
     croak("Unknown property '%s' passed to setloginproperty", prop_name);
   }


   // Some properties affects others.
   if (gbl_init_props[ix].propset_enum == oleinit_props &&
       gbl_init_props[ix].property_id == DBPROP_AUTH_USERID) {
      // If userid is set, we clear Integrated security.
      setloginproperty(olle_ptr, "IntegratedSecurity", &PL_sv_undef);
   }
   else if (gbl_init_props[ix].propset_enum == oleinit_props &&
            gbl_init_props[ix].property_id == DBPROP_INIT_PROVIDERSTRING) {
      // In this case, all other properties should be flushed.
      for (int j = 0; gbl_init_props[j].propset_enum != datasrc_props; j++) {
         VariantClear(&mydata->init_properties[j].vValue);
      }
   }

   // If the server changes, the SQL_version attribute is no longer valid.
   // We cannot set it to undef - Perl moans about read-only. But delete works!
   if (gbl_init_props[ix].propset_enum == oleinit_props &&
       (gbl_init_props[ix].property_id == DBPROP_INIT_PROVIDERSTRING ||
        gbl_init_props[ix].property_id == DBPROP_INIT_DATASOURCE ||
        gbl_init_props[ix].property_id == SSPROP_INIT_NETWORKADDRESS)) {
      delete_from_hash(olle_ptr, HV_SQLversion);
   }

   // First clear the current value and set property to VT_EMPTY.
   VariantClear(&mydata->init_properties[ix].vValue);

   // Then set the value appropriately
   if (prop_value && SvOK(prop_value)) {
      mydata->init_properties[ix].vValue.vt = gbl_init_props[ix].datatype;

      // First handle any specials. Currently there are two.
      if (gbl_init_props[ix].propset_enum == oleinit_props &&
          gbl_init_props[ix].property_id == DBPROP_INIT_OLEDBSERVICES) {
         // For OLE DB Services, we are only using connection pooling.
         mydata->init_properties[ix].vValue.lVal = (SvTRUE(prop_value)
                 ? DBPROPVAL_OS_RESOURCEPOOLING : DBPROPVAL_OS_DISABLEALL);
      }
      else if (gbl_init_props[ix].propset_enum == oleinit_props &&
               gbl_init_props[ix].property_id == DBPROP_AUTH_INTEGRATED) {
         // For integrated security, handle numeric values gently.
         if (SvIOK(prop_value)) {
            if (SvIV(prop_value) != 0) {
                mydata->init_properties[ix].vValue.bstrVal =
                    SysAllocString(L"SSPI");
            }
            else {
               mydata->init_properties[ix].vValue.vt = VT_EMPTY;
             }
         }
         else {
            mydata->init_properties[ix].vValue.bstrVal = SV_to_BSTR(prop_value);
         }
      }
      else {
         switch(gbl_init_props[ix].datatype) {
            case VT_BOOL :
                mydata->init_properties[ix].vValue.boolVal =
                    (SvTRUE(prop_value) ? VARIANT_TRUE : VARIANT_FALSE);
                break;

            case VT_I2 :
                mydata->init_properties[ix].vValue.iVal = (SHORT) SvIV(prop_value);
                break;

            case VT_UI2 :
                mydata->init_properties[ix].vValue.uiVal = (USHORT) SvIV(prop_value);
                break;

            case VT_I4 :
                mydata->init_properties[ix].vValue.lVal = SvIV(prop_value);
                break;

            case VT_BSTR :
                mydata->init_properties[ix].vValue.bstrVal = SV_to_BSTR(prop_value);
                break;

            default :
               croak ("Internal error: Unexpected datatype %d when setting property '%s'",
                      gbl_init_props[ix].datatype, prop_name);
          }
       }
   }
}

void disconnect(SV * olle_ptr)
{
    internaldata * mydata = get_internaldata(olle_ptr);

    free_batch_data(mydata);
    free_ole_ptr(mydata->datasrc_ptr);  // This disconnects - or returns to pool.
    free_ole_ptr(mydata->init_ptr);
}

//====================================================================
// Utility routines.
//====================================================================

//---------------------------------------------------------------------
// This is a helper routine called from get_object_id in the Perl code.
// It cracks an object specification into its parts, retaining any quotes
// around the identifiers and returns the result in sv_server, sv_db,
// sv_schema and sv_object.
//------------------------------------------------------------------
static void parsename(SV   * olle_ptr,
                      SV   * sv_namestr,
                      int    retain_quotes,
                      SV   * sv_server,
                      SV   * sv_db,
                      SV   * sv_schema,
                      SV   * sv_object)
{
   STRLEN namelen;
   char  * namestr = SvPV(sv_namestr, namelen);
   char  * server = NULL;
   char  * db = NULL;
   char  * schema = NULL;
   char  * object = NULL;
   STRLEN  inix = 0;
   STRLEN  outix = 0;
   int     dotno = 0;
   char    endtoken = '\0';
   BOOL    lastwasendtoken = FALSE;

   New(902, object, namelen + 1, char);
   memset(object, 0, namelen + 1);
   outix = 0;

   while (inix < namelen) {
      char chr = namestr[inix++];

      if (outix == 0  && ! endtoken) {
         // We are at the first character in an element. Only here a quote
         // delimiter is legal.
         endtoken = '\0';
         if (chr == '"')
            endtoken = '"';
         if (chr == '[')
            endtoken = ']';
         if (endtoken && ! retain_quotes)
            continue;
      }

      if (endtoken) {
         if (retain_quotes) {
            // We are in a quoted element, and with retain_quotes, we should
            // in this case always copy to the output, which always is
            // object (could be moved, see below.)
               object[outix++] = chr;

            // Check if we are at end of the delimiter. Note that if outix == 1
            // we are looking at the opening quote character.
            if (outix > 1 && chr == endtoken && namestr[inix] != endtoken) {
               endtoken = '\0';
            }
         }
         else {
            // If we are not retaining quotes, we should only copy an endtoken
            // if previous was also an end token. And never the first character!
            if (chr != endtoken || (chr == endtoken && lastwasendtoken)) {
               object[outix++] = chr;
            }
            else if (chr == endtoken && namestr[inix] != endtoken) {
            // Else if we have an endtoken and next character is not one, we
            // are at the end.
               endtoken = '\0';
            }

            // Set lastwasendtoken. It never stays true very long.
            lastwasendtoken = (lastwasendtoken ? FALSE : (chr == endtoken));
         }

      }
      else {
         switch (chr) {
            case ' '  :
            case '\t' :
            case '\n' :
               // White-space. Ignore.
               break;

            case '.' : {
               // Found a dot. Move what we save in object to schema, and
               // schema to db if this was the second dot.
               dotno++;
               switch (dotno) {
                  case 1 : schema = object;
                           break;

                  case 2 : db = schema;
                           schema = object;
                           break;

                  case 3 : server = db;
                           db = schema;
                           schema = object;
                           break;

                  default :
                     // Too many dots, just copy.
                     object[outix++] = chr;
                     break;
               }

               // Allocate new buffer.
               New(902, object, namelen + 1, char);
               memset(object, 0, namelen + 1);
               outix = 0;

               break;
            }

            default :
              // Plain copy.
              object[outix++] = chr;
              break;
         }
      }
   }

   if (endtoken) {
      // Input string is terminated, but the identifier was not closed. Cry foul.
      olle_croak(olle_ptr, "Object specification '%s' has an unterminated quoted identifier",
                 namestr);
   }

   // Set output parameters.
   if (server) {
      sv_setpvn(sv_server, server, strlen(server));
      if (SvUTF8(sv_namestr)) {
          SvUTF8_on(sv_server);
      }
      else {
          SvUTF8_off(sv_server);
      }
      Safefree(server);
   }
   else {
      sv_setpvn(sv_server, "", 0);
   }


   if (db) {
      sv_setpvn(sv_db, db, strlen(db));
      if (SvUTF8(sv_namestr)) {
          SvUTF8_on(sv_db);
      }
      else {
          SvUTF8_off(sv_db);
      }
      Safefree(db);
   }
   else {
      sv_setpvn(sv_db, "", 0);
   }

   if (schema) {
      sv_setpvn(sv_schema, schema, strlen(schema));
      if (SvUTF8(sv_namestr)) {
          SvUTF8_on(sv_schema);
      }
      else {
          SvUTF8_off(sv_schema);
      }
      Safefree(schema);
   }
   else {
      sv_setpvn(sv_schema, "", 0);
   }

   sv_setpvn(sv_object, object, strlen(object));
   if (SvUTF8(sv_namestr)) {
       SvUTF8_on(sv_object);
   }
   else {
       SvUTF8_off(sv_object);
   }
   Safefree(object);
}

//----------------------------------------------------------------------
// This is a helper routine that scans a SQL commad string for ? and
// replaces them with @P1, @P2 etc. It's called from setup_sqlcommand
// in the Perl module, and not called by the XS code. It's written in C++,
// simply because it appeared simpler than to do it in Perl.
//----------------------------------------------------------------------
void
replaceparamholders (SV * olle_ptr,
                     SV * cmdstring)
{
   STRLEN inputlen;
   char * inputorg = SvPV(cmdstring, inputlen);
   char * input;
   char * output;
   STRLEN inix  = 0;
   STRLEN outix = 0;
   int    parno = 1;
   char   paramstr[12];
   char   endtoken = '\0';
   int    cmtnestlvl = 0;

   // Since we do some lookahead, we copy the string a buffer which is
   // somewhat larger, so we are not looking at someone else's memory.
   New(902, input, inputlen + 3, char);
   memcpy(input, inputorg, inputlen);
   input[inputlen ] = ' ';
   input[inputlen + 1] = ' ';
   input[inputlen + 2] = ' ';

   // The output buffer we make three times as large, since a ? gets
   // replaced with at least three chars.
   New(902, output, 3*inputlen + 3, char);

   // Yeah, the condition is such that in some weird cases, we do not copy
   // all characters. We don't expect this to occur in the real world.
   while (inix < inputlen && outix < 3*inputlen - 3) {
      char chr = input[inix++];

      if (! endtoken) {
      // We are in regular code - not a comment, string lit or quoted identifier.
         if (chr == '?') {
         // Expand ? to @p1 etc.
            sprintf(paramstr, "@P%d", parno++);
            strcpy(&(output[outix]), paramstr);
            outix += strlen(paramstr);
         }
         else {
         // Copy the character as is, and look for start of comment or string.
            output[outix++] = chr;
            switch (chr) {
               case '/' : if (input[inix] == '*') {
                          // Note that /* can nest.
                             endtoken = '/';
                             cmtnestlvl++;

                             // Must move on two chars, or else /*/ would
                             // be both start and end of comment.
                             output[outix++] = input[inix++];
                             output[outix++] = input[inix++];
                          }
                          break;
               case '-' : if (input[inix] == '-') {
                             endtoken = '\n';
                          }
                          break;
               case '\'' : endtoken = '\'';
                           break;
               case '"'  : endtoken = '"';
                           break;
               case '['  : endtoken = ']';
                           break;
            }
         }
      }
      else {
      // We are in some special state. Copy character, no ?-expanding here.
         output[outix++] = chr;
         if (chr == endtoken) {
            switch (chr) {
               case '/'  : if (input[inix - 2] == '*') {
                           // Lookback to see if we have a */, note that they
                           // can nest.
                              cmtnestlvl--;
                              if (! cmtnestlvl) {
                                 endtoken = '\0';
                              }
                           }
                           else if (input[inix] == '*') {
                           // Nested comment. Again we must move on two chars.
                              cmtnestlvl++;
                              output[outix++] = input[inix++];
                              output[outix++] = input[inix++];
                           }
                           break;
              case  '\n' : endtoken = '\0';
                           break;
              case  '\'' :
              case  '"'  :
              case  ']'  : // If doubled, this is a false alarm. Copy the
                           // double now, and move on.
                           if (input[inix] == endtoken) {
                              output[outix++] = input[inix++];
                           }
                           else {
                              endtoken = '\0';
                           }
                           break;
            }
         }
      }
   }


   sv_setpvn(cmdstring, output, outix);
   Safefree(output);
   Safefree(input);
}

//--------------------------------------------------------------------
// This helper routine extracts the encoding for an XML string, and
// deduces whether it is a 16/32-bit encoding, UTF-8 or a 8-bit charset.
// It also returns the position where the encoding name appears, or -1
// if there isn't any.
//---------------------------------------------------------------------
typedef enum xmlcharsettypes {eightbit, utf8, sixteen} xmlcharsettypes;
void get_xmlencoding (SV              * sv,
                      xmlcharsettypes &xmlcharsettype,
                      int             &charsetpos)
{
   char   encoding[20];
   int    scanret;
   char * str;

   if (sv == NULL || ! SvOK(sv)) {
      xmlcharsettype = utf8;
      charsetpos    = -1;
      return;
   }
   str = SvPV_nolen(sv);

   // If there is an encoding, it must come in the prolog which must be at
   // the very first in the file. This is heaviliy regimented by the XML
   // standard. sscanf comes in handy here.
   scanret = sscanf(str, "<?xml version = \"1.0\" encoding = \"%h19[^\"]\"",
                    encoding);

   // scanret == 1 => we found an encoding string.
   if (scanret) {
      // Get the position.
      char *tmp = strstr(str, encoding);
      charsetpos = tmp - str;

      // Then normalise to lowercase.
      _strlwr(encoding);

      // Then compare to various known encodings.
      if (strstr(encoding, "utf-8") == encoding) {
         xmlcharsettype = utf8;
      }
      else if (strstr(encoding, "ucs-2") == encoding ||
              strstr(encoding, "utf-16") == encoding) {
         xmlcharsettype = sixteen;
      }
      else {
         // All other encodings are assumed to be 8-bit.
         xmlcharsettype = eightbit;
      }
   }
   else {
      // If there was no encoding, then it has to be UTF-8.
      xmlcharsettype = utf8;
      charsetpos     = -1;
   }
}

//==================================================================
// Get data (and command) INTO SQL Server. Conversion from SV to SQL
// types first, and then $X->initbatch, $X->enterparameter and
// $X->executebatch.
//==================================================================

//------------------------------------------------------------------
// Conversion-from-SV routines. These routines converts an SV to the
// desired SQL Server type. For most types the conversion is implicit
// from the data type of the Perl variable.
// Note that SV_to_BSTR is in the beginning of the file, as this is a
// generally used routine.
//------------------------------------------------------------------

// This is a helper routine, which uses DataConvert to convert a Perl
// String, which can be either an 8-bit string or a UTF8-string.
HRESULT  SVstr_to_sqlvalue (SV   * sv,
                            DBTYPE sqltype,
                            void * sqlvalue,
                            BYTE   precision = NULL,
                            BYTE   scale     = NULL)
{
   HRESULT ret;

   assert(SvPOK(sv));
   if (SvUTF8(sv)) {
      DBLENGTH bytelen;
      BSTR bstr = SV_to_BSTR(sv, &bytelen);
      ret = data_convert_ptr->DataConvert(
            DBTYPE_WSTR, sqltype, bytelen, NULL,
            bstr, sqlvalue, NULL, DBSTATUS_S_OK, NULL,
            precision, scale, 0);
      SysFreeString(bstr);
   }
   else {
      STRLEN strlen;
      char * str = SvPV(sv, strlen);
      ret = data_convert_ptr->DataConvert(
            DBTYPE_STR, sqltype, strlen, NULL,
            str, sqlvalue, NULL, DBSTATUS_S_OK, NULL,
            precision, scale, 0);
   }

   return ret;
}


BOOL SV_to_bigint (SV      * sv,
                   LONGLONG  &bigintval)
{
   HRESULT ret;

   if (SvPOK(sv)) {
      HRESULT ret = SVstr_to_sqlvalue(sv, DBTYPE_I8, &bigintval);
   }
   else if (SvNOK(sv)) {
      double dbl = SvNV(sv);
      ret = data_convert_ptr->DataConvert(
            DBTYPE_R8, DBTYPE_I8, sizeof(double), NULL,
            &dbl, &bigintval, NULL, DBSTATUS_S_OK, NULL,
            0, 0, 0);
   }
   else {
   // It could be an integer or a reference, whatever we handle as int.
      bigintval = SvIV(sv);
      ret = S_OK;
   }

   return SUCCEEDED(ret);
}

BOOL SV_to_binary (SV        * sv,
                   bin_options optBinaryAsStr,
                   BOOL        istimestamp,
                   BYTE      * &binaryval,
                   DBLENGTH    &value_len)
{
    BOOL     retval;
    STRLEN   perl_len;
    char   * perl_ptr = (char *) SvPV(sv, perl_len);

    if (optBinaryAsStr != bin_binary) {
       HRESULT  ret;

       // Note that we don't here consider the possibility that the string
       // may be a UTF-8 string. It should really only include 0-9 and
       // A-F plus any leading 0x. Digits from other scripts are not
       // considered.

       if (_strnicmp(perl_ptr, "0x", 2) == 0) {
          perl_ptr += 2;
          perl_len -= 2;
       }

       value_len = perl_len / 2;
       New(902, binaryval, value_len, BYTE);
       ret = data_convert_ptr->DataConvert(
             DBTYPE_STR, DBTYPE_BYTES, perl_len, NULL,
             perl_ptr, binaryval, value_len, DBSTATUS_S_OK, NULL,
             NULL, NULL, 0);
       retval = SUCCEEDED(ret);
    }
    else {
       value_len = perl_len;
       New(902, binaryval, value_len, BYTE);
       memcpy(binaryval, perl_ptr, value_len);
       retval = TRUE;
    }

    // If this is a timestamp value, and the input value gave us a value
    // less than 8 bytes, we must reallocate and pad, since timestamp is
    // is fixed-length and else we would send random garbage.
    if (istimestamp && value_len < 8) {
       BYTE *tmp;
       New(902, tmp, 8, BYTE);
       memset(tmp, 0, sizeof(BYTE) * 8);
       memcpy(tmp, binaryval, value_len);
       value_len = 8;
       Safefree(binaryval);
       binaryval = tmp;
    }

    return retval;
}

BOOL SV_to_char (SV       * sv,
                 char     * &charval,
                 DBLENGTH   &value_len)
{
   if (! SvUTF8(sv)) {
      // This is quite trivial. We should however copy the string to our own
      // buffer.
      STRLEN strlen;
      char * perl_str = SvPV(sv, strlen);
      value_len = strlen;
      New(902, charval, strlen + 1, char);
      memcpy(charval, perl_str, strlen);
      return TRUE;
   }
   else {
      // Use IDataConvert to get a string in ANSI CP.
      DBLENGTH bytelen;
      BSTR     bstr = SV_to_BSTR(sv, &bytelen);
      HRESULT  ret;

      New(902, charval, bytelen / 2 + 1, char);
      value_len = bytelen / 2;

      ret = data_convert_ptr->DataConvert(
            DBTYPE_WSTR, DBTYPE_STR, bytelen, NULL,
            bstr, charval, (bytelen / 2 + 1), DBSTATUS_S_OK, NULL,
            NULL, NULL, 0);
      SysFreeString(bstr);
      return SUCCEEDED(ret);
   }
}

BOOL SV_to_XML (SV        * sv,
                BOOL        &is_8bit,
                char      * &xmlchar,
                BSTR        &xmlbstr,
                DBLENGTH    &value_len)
{
   xmlcharsettypes   charsettype;
   int               dummy;
   BOOL              retval;

   // Get the character-set type.
   get_xmlencoding(sv, charsettype, dummy);

   // And then handle the string accordingly.
   switch (charsettype) {
      case eightbit :
         retval = SV_to_char(sv, xmlchar, value_len);
         //value_len *= 2;
         is_8bit = TRUE;
         xmlbstr = NULL;
         break;

      case utf8 : {
         // Force string to be UTF-8.
         STRLEN strlen;
         char   * perl_str = SvPVutf8(sv, strlen);
         value_len = strlen;
         New(902, xmlchar, strlen + 1, char);
         memcpy(xmlchar, perl_str, strlen);
         is_8bit = true;
         xmlbstr = NULL;
         retval = TRUE;
         break;
      }

      case sixteen : {
         // Convert to BSTR and force insert of a BOM.
         xmlbstr = SV_to_BSTR(sv, &value_len, TRUE);
         xmlchar = NULL;
         is_8bit = FALSE;
         retval = TRUE;
         break;
      }
      default :
         croak ("Entirely unexpected value for charsettype %d", charsettype);
         break;
   }
   return retval;
}

// This is a helper routine to SV_to_datetime.
BOOL get_datetime_hashvalue(SV         * olle_ptr,
                            HV         * hv,
                            const char * part,
                            BOOL         mandatory,
                            LONG         minval,
                            LONG         maxval,
                            SHORT       &partval)
{
   SV    ** svp;
   SV    *  sv = NULL;
   LONG     intvalue;

   partval = 0;
   svp = hv_fetch(hv, part, strlen(part), 0);
   if (svp != NULL)
       sv = *svp;
   if (sv == NULL || ! SvOK(sv)) {
      if (! mandatory) {
         return TRUE;
      }
      else {
         BSTR msg = SysAllocStringLen(NULL, 200);
         wsprintf(msg, L"Mandatory part '%S' missing from datetime hash.", part);
         olledb_message(olle_ptr, -1, 1, 10, msg);
         SysFreeString(msg);
         return FALSE;
      }
   }
   intvalue = SvIV(sv);
   if (intvalue < minval || intvalue > maxval) {
      BSTR msg = SysAllocStringLen(NULL, 200);
      wsprintf(msg, L"Part '%S' in dateiume hash has illegal value %d.", part, intvalue);
      olledb_message(olle_ptr, -1, 1, 10, msg);
      SysFreeString(msg);
      return FALSE;
   }
   partval = (SHORT) intvalue;
   return TRUE;
}

BOOL SV_to_datetime (SV          * sv,
                     DBTIMESTAMP &datetime,
                     SV          * olle_ptr)
{
   if (SvROK(sv)) {
      HV    * hv;
      SHORT partvalue;

      // It is a reference, but it is a hash reference?
      if (strncmp(SvPV_nolen(sv), "HASH(", 5) != 0)
         return FALSE;

      hv = (HV *) SvRV(sv);

      if (! get_datetime_hashvalue(olle_ptr, hv, "Year", TRUE, 0, 9999, partvalue))
         return FALSE;
      datetime.year = (SHORT) partvalue;

      if (! get_datetime_hashvalue(olle_ptr, hv, "Month", TRUE, 1, 12, partvalue))
         return FALSE;
      datetime.month = (USHORT) partvalue;

      if (! get_datetime_hashvalue(olle_ptr, hv, "Day", TRUE, 1, 31, partvalue))
         return FALSE;
      datetime.day = (USHORT) partvalue;

      if (! get_datetime_hashvalue(olle_ptr, hv, "Hour", FALSE, 0, 24, partvalue))
         return FALSE;
      datetime.hour = (USHORT) partvalue;

      if (! get_datetime_hashvalue(olle_ptr, hv, "Minute", FALSE, 0, 59, partvalue))
         return FALSE;
      datetime.minute = (USHORT) partvalue;

      if (! get_datetime_hashvalue(olle_ptr, hv, "Second", FALSE, 0, 61, partvalue))
         return FALSE;
      datetime.second = (SHORT) partvalue;

      if (! get_datetime_hashvalue(olle_ptr, hv, "Fraction", FALSE, 0, 999, partvalue))
         return FALSE;
      datetime.fraction = partvalue * 1000000;

      return TRUE;
   }
   else if (SvNOK(sv) || SvIOK(sv)) {
      DATE           dateval = SvNV(sv);
      HRESULT        ret;

      ret = data_convert_ptr->DataConvert(
            DBTYPE_DATE, DBTYPE_DBTIMESTAMP, sizeof(DATE),
            NULL, &dateval, &datetime, NULL, DBSTATUS_S_OK, NULL,
            NULL, NULL, 0);

      return (SUCCEEDED(ret));
   }
   else {
      // Looks like it is a string. At least we treat it as such.
      STRLEN    perl_len;
      char    * perl_str = SvPV(sv, perl_len);
      char    * str;
      char    * p;
      DBLENGTH  strlen = perl_len;
      BOOL      done_ownmod = FALSE;
      BOOL      have_DATEval = FALSE;
      DATE      dateval;
      HRESULT   ret;

      // This is a little messy. We want to support strings in UTF-8 format
      // (may be month names), but we also want to support the YYYYMMDD format
      // and XML format, which IDataConvert does not support, so we need to
      // detect these and modify the string in this case. The good news is
      // if we make these modifications, we can ignore UTF-8.

      // Set up string that is a copy of the Perl string.
      New(902, str, perl_len + 10, char);
      p = str;

      // In this loop we insert - in YYYYMMDD to make it YYYY-MM-DD.
      if (perl_len >= 8) {
         BOOL   seen_nondigit = FALSE;
         for (int i = 1; i <= 8; i++) {
            seen_nondigit = (seen_nondigit || ! isdigit(* perl_str));
            if (! seen_nondigit && (i == 5 || i == 7) && isdigit(* perl_str)) {
               *p++ = '-';
               strlen++;
               done_ownmod = TRUE;
            }
            *p = *perl_str;
            p++;
            perl_str++;
            perl_len--;
         }
      }
      // Then copy the rest of the string.
      memcpy(p, perl_str, perl_len);

      // Next task is to replace a T or Z in position 11, but only if we now
      // have YYYY-MM-DDTHH:MM or YYYY-MM-DDZ. For T we don't check the time,
      // part, but we don't remove the T if length is too short. For Z
      // there must be no trailing chars.
      if (strlen >= 11) {
         p = str;
         BOOL couldbeansi = TRUE;
         for (int i = 1; i <= 10; i++) {
            couldbeansi = (couldbeansi &&
                           ((i == 5 || i == 8) ? (*p == '-') : isdigit(*p)));
            p++;
         }
         if (couldbeansi) {
            if (*p == 'T' && strlen >= 15) {
               *p = ' ';
               done_ownmod = TRUE;
            }
            if (*p == 'Z' && strlen == 11) {
               *p = '\0';
               done_ownmod = TRUE;
            }
         }
      }

      if (done_ownmod || ! SvUTF8(sv)) {
      // Main track: string is not UTF-8, or we have modified it. Note that
      // even if we did not modify str, we don't have to use it.
         // First we try to convert it as an ISO string.
         ret = data_convert_ptr->DataConvert(
               DBTYPE_STR, DBTYPE_DBTIMESTAMP, strlen, NULL,
               str, &datetime, NULL, DBSTATUS_S_OK, NULL,
               NULL, NULL, 0);

         if (FAILED(ret)) {
            // ISO format failed. Try regional settings. This is a two-step
            // process, by going over DATE which is a float value.
            have_DATEval = TRUE;
            ret = data_convert_ptr->DataConvert(
                  DBTYPE_STR, DBTYPE_DATE, strlen, NULL,
                  str, &dateval, NULL, DBSTATUS_S_OK, NULL,
                  NULL, NULL, 0);
         }
      }
      else {
      // String is UTF-8. Switch to UTF-16 and work from there.
         DBLENGTH  bytelen;
         BSTR      bstr = SV_to_BSTR(sv, &bytelen);

         // Try ISO string...
         ret = data_convert_ptr->DataConvert(
               DBTYPE_WSTR, DBTYPE_DBTIMESTAMP, bytelen, NULL,
               bstr, &datetime, NULL, DBSTATUS_S_OK, NULL,
               NULL, NULL, 0);

         if (FAILED(ret)) {
            // And then regional settings.
            have_DATEval = TRUE;
            ret = data_convert_ptr->DataConvert(
                  DBTYPE_WSTR, DBTYPE_DATE, bytelen, NULL,
                  bstr, &dateval, NULL, DBSTATUS_S_OK, NULL,
                  NULL, NULL, 0);
         }
         SysFreeString(bstr);
      }

      Safefree(str);   // Not needed any more.

      // If we have a DATE value at this point, we are half-way through an
      // attempt to use regional settings - which may have failed. If we
      // don't have a DATE value, we are fine.
      if (have_DATEval) {
         if (FAILED(ret)) {
            return FALSE;
         }

         // One would expect that this last step always suceedes.
         ret = data_convert_ptr->DataConvert(
               DBTYPE_DATE, DBTYPE_DBTIMESTAMP, sizeof(DATE), NULL,
               &dateval, &datetime, NULL, DBSTATUS_S_OK, NULL,
               NULL, NULL, 0);
         return TRUE;
      }
      else {
         return TRUE;
      }
   }
}

BOOL SV_to_decimal(SV        * sv,
                   BYTE        precision,
                   BYTE        scale,
                   DB_NUMERIC &decimalval)
{
   HRESULT  ret;

   if (SvPOK(sv)) {
      ret = SVstr_to_sqlvalue(sv, DBTYPE_NUMERIC, &decimalval,
                              precision, scale);
   }
   else {
      double dbl = SvNV(sv);
      ret = data_convert_ptr->DataConvert(
            DBTYPE_R8, DBTYPE_NUMERIC, sizeof(double), NULL,
            &dbl, &decimalval, NULL, DBSTATUS_S_OK, NULL,
            precision, scale, 0);
   }
   return SUCCEEDED(ret);
}

BOOL SV_to_GUID (SV       * sv,
                 GUID       &guidval)
{
   if (SvPOK(sv)) {
      HRESULT ret;
      STRLEN strlen;
      char * perl_str = SvPV(sv, strlen);

      if (strlen == 36) {
         // This could be a GUID without braces, so we add them.
         char guidstr[39];
         sprintf(guidstr, "{%s}", perl_str);
         ret = data_convert_ptr->DataConvert(
               DBTYPE_STR, DBTYPE_GUID, 38, NULL,
               guidstr, &guidval, NULL, DBSTATUS_S_OK, NULL,
               NULL, NULL, 0);
      }
      else {
         ret = SVstr_to_sqlvalue(sv, DBTYPE_GUID, &guidval);
      }
      return SUCCEEDED(ret);
   }
   else {
      // It would be useless even to try...
      return FALSE;
   }
}

BOOL SV_to_money(SV * sv,
                 CY  &moneyval)
{
   HRESULT  ret;

   if (SvPOK(sv)) {
      ret = SVstr_to_sqlvalue(sv, DBTYPE_CY, &moneyval);
   }
   else {
      double dbl = SvNV(sv);
      ret = data_convert_ptr->DataConvert(
            DBTYPE_R8, DBTYPE_CY, sizeof(double), NULL,
            &dbl, &moneyval, sizeof(CY), DBSTATUS_S_OK, NULL,
            NULL, NULL, 0);
   }
   return SUCCEEDED(ret);
}

BOOL SV_to_ssvariant (SV        * sv,
                      SSVARIANT   &variant,
                      SV        * olle_ptr,
                      void      * &save_str,
                      BSTR        &save_bstr)
{
    save_str = NULL;
    save_bstr = NULL;
    memset(&variant, 0, sizeof(SSVARIANT));

    // If the SV is a reference to a hash, it may be a datetime value, so
    // we try this first.
    if (SvROK(sv)) {
       DBTIMESTAMP dateval;
       if (SV_to_datetime(sv, dateval, olle_ptr)) {
          variant.vt = VT_SS_DATETIME;
          variant.tsDateTimeVal = dateval;
          return TRUE;   // Must exit here, since there is a fall-through.
       }
    }

    if (SvIOK(sv)) {
       variant.vt = VT_SS_I4;
       variant.lIntVal = SvIV(sv);
    }
    else if (SvNOK(sv)) {
       variant.vt = VT_SS_R8;
       variant.dblFloatVal = SvNV(sv);
    }
    else if (SvUTF8(sv)) {
       DBLENGTH bytelen;
       BSTR     bstr = SV_to_BSTR(sv, &bytelen);

       if (bytelen > 8000) bytelen = 8000;
       variant.vt = VT_SS_WVARSTRING;
       variant.NCharVal.sActualLength = (SHORT) bytelen;
       variant.NCharVal.sMaxLength = (SHORT) bytelen;
       variant.NCharVal.pwchNCharVal = bstr;
       save_bstr = bstr;
    }
    else {
    // That is, we end up here, even if the value is a reference or whatever.
       STRLEN strlen;
       char * perl_ptr = SvPV(sv, strlen);
       char * str;

       if (strlen > 8000) strlen = 8000;
       New(902, str, strlen + 1, char);
       memcpy(str, perl_ptr, strlen);
       str[strlen] = '\0';
       variant.vt = VT_SS_VARSTRING;
       variant.CharVal.sActualLength = strlen;
       variant.CharVal.sMaxLength = strlen;
       variant.CharVal.pchCharVal = str;
       save_str = str;
    }

    return TRUE;
}

// This is a different SV_to_xxx thing. For UDT and XML, there may be
// parameter properties to add to the parameter record.
static void add_param_props (SV        * olle_ptr,
                             paramdata * param,
                             SV        * typeinfo)
{
    // Drop out if there is no typeinfo.
    if (! typeinfo || ! SvOK(typeinfo)) {
       return;
    }

    SV * server   = newSV(sv_len(typeinfo));
    SV * database = newSV(sv_len(typeinfo));
    SV * schema   = newSV(sv_len(typeinfo));
    SV * object   = newSV(sv_len(typeinfo));
    int  ix = 0;

    // First extract components from typeinfo.
    parsename(olle_ptr, typeinfo, 0, server, database, schema, object);

    // If there was a server, cry foul.
    if (sv_len(server) > 0) {
       BSTR typeinfo_str = SV_to_BSTR(typeinfo);
       BSTR msg = SysAllocStringLen(NULL, SysStringLen(typeinfo_str) + 200);
       SysFreeString(typeinfo_str);
       wsprintf(msg, L"Type name/XML schema '%s' includes a server compenent.\n",
                typeinfo_str);
       olledb_message(olle_ptr, -1, -1, 16, msg);
       SysFreeString(msg);
       SvREFCNT_dec(server);
       SvREFCNT_dec(database);
       SvREFCNT_dec(schema);
       SvREFCNT_dec(object);
       return;
    }

    // Find out how many components we have.
    if (sv_len(database) > 0) param->param_props_cnt++;
    if (sv_len(schema) > 0) param->param_props_cnt++;
    if (sv_len(object) > 0) param->param_props_cnt++;

    // If there was nothing, just drop out.
    if (param->param_props_cnt == 0)
        return;

    // Now we can allocate as many properties as need
    New(902, param->param_props, param->param_props_cnt, DBPROP);

    // Store server if any.
    if (sv_len(database) > 0) {
       param->param_props[ix].dwPropertyID =
            (param->datatype == DBTYPE_UDT ?
             SSPROP_PARAM_UDT_CATALOGNAME :
             SSPROP_PARAM_XML_SCHEMACOLLECTION_CATALOGNAME);
       param->param_props[ix].colid = DB_NULLID;
       param->param_props[ix].dwOptions = DBPROPOPTIONS_REQUIRED;
       VariantInit(&(param->param_props[ix].vValue));
       param->param_props[ix].vValue.vt = VT_BSTR;
       param->param_props[ix].vValue.bstrVal = SV_to_BSTR(database);
       ix++;
    }

    // And schema if any.
    if (sv_len(schema) > 0) {
       param->param_props[ix].dwPropertyID =
           (param->datatype == DBTYPE_UDT ?
            SSPROP_PARAM_UDT_SCHEMANAME :
            SSPROP_PARAM_XML_SCHEMACOLLECTION_SCHEMANAME);
       param->param_props[ix].colid = DB_NULLID;
       param->param_props[ix].dwOptions = DBPROPOPTIONS_REQUIRED;
       VariantInit(&(param->param_props[ix].vValue));
       param->param_props[ix].vValue.vt = VT_BSTR;
       param->param_props[ix].vValue.bstrVal = SV_to_BSTR(schema);
       ix++;
    }

    // And the type name.
    // Store server if any.
    if (sv_len(object) > 0) {
       param->param_props[ix].dwPropertyID =
            (param->datatype == DBTYPE_UDT ?
             SSPROP_PARAM_UDT_NAME :
             SSPROP_PARAM_XML_SCHEMACOLLECTIONNAME);
       param->param_props[ix].colid = DB_NULLID;
       param->param_props[ix].dwOptions = DBPROPOPTIONS_REQUIRED;
       VariantInit(&(param->param_props[ix].vValue));
       param->param_props[ix].vValue.vt = VT_BSTR;
       param->param_props[ix].vValue.bstrVal = SV_to_BSTR(object);
    }

    // We must clean up our SVs to not leak memory.
    SvREFCNT_dec(server);
    SvREFCNT_dec(database);
    SvREFCNT_dec(schema);
    SvREFCNT_dec(object);
}

//--------------------------------------------------------------------
// $X->initbatch.
//--------------------------------------------------------------------
void initbatch(SV * olle_ptr,
               SV * sv_cmdtext)
{
    internaldata       * mydata = get_internaldata(olle_ptr);

    if (! (sv_cmdtext && SvOK(sv_cmdtext))) {
       olle_croak(olle_ptr, "Parameter sv_cmdtext to submitcmd missing or is undef");
    }

    // There must be no pending command, as then a command is still progress.
    if (mydata->pending_cmd != NULL) {
        olle_croak(olle_ptr, "Cannot init a new batch, when previous batch has not been processed");
    }

    // Save the command.
    mydata->pending_cmd = SV_to_BSTR(sv_cmdtext);
}

//------------------------------------------------------------------------
// $X->enterparameter
//------------------------------------------------------------------------
int enterparameter(SV   * olle_ptr,
                   SV   * sv_nameoftype,
                   SV   * sv_maxlen,
                   SV   * paramname,
                   BOOL   isinput,
                   BOOL   isoutput,
                   SV   * sv_value,
                   BYTE   precision,
                   BYTE   scale,
                   SV   * typeinfo)
{
   internaldata    * mydata = get_internaldata(olle_ptr);
   ULONG             maxlen;
   char            * nameoftype;
   paramdata       * this_param;
   DBBINDING       * binding;     // Shortcut to this_param->binding.
   DBPARAMBINDINFO * param_info;  // Shortcut to this_param->param_info.
   paramvalue      * value;       // Shortcur to this_param->value.
   BOOL              istimestamp = FALSE;
   BOOL              value_OK = TRUE;


   // Check that we're in the state where we're accepting parameters.
   if (mydata->pending_cmd == NULL) {
      olle_croak(olle_ptr, "Cannot call enterparameter now. There is a pending command. Call initbatch first");
   }

   if (mydata->cmdtext_ptr != NULL) {
      olle_croak(olle_ptr, "Cannot call enterparameter now. There are unprocessed resultsets. Call cancelbatch first");
   }

   // Type name is mandatory.
   if (! sv_nameoftype || ! SvOK(sv_nameoftype)) {
      olle_croak(olle_ptr, "You must pass a legal type name to enterparameter. Cannot pass undef");
   }
   nameoftype = SvPV_nolen(sv_nameoftype);

   // Get maxlen.
   if (sv_maxlen && SvOK(sv_maxlen)) {
      maxlen = SvIV(sv_maxlen);
   }
   else {
      maxlen = 0;
   }

   // Allocate space for this parameter.
   New(902, this_param, 1, paramdata);
   memset(this_param, 0, sizeof(paramdata));

   // Fill it in.
   this_param->datatype = lookup_type_map(nameoftype);

   // Timestamp requires special precautions, but looks just like any other
   // binary value
   if (this_param->datatype == DBTYPE_BYTES &&
       (strcmp(nameoftype, "timestamp") == 0 ||
        strcmp(nameoftype, "rowversion") == 0)) {
       istimestamp = TRUE;
   }

   // input/output maps to flags.
   this_param->isinput = isinput;
   this_param->isoutput = isoutput;

   // Is value NULL or not?
   this_param->isnull = (! isinput || ! sv_value || ! SvOK(sv_value));

   // Link in to the parameter list and increase parameter count.
   this_param->next = NULL;
   if (mydata->paramlast == NULL) {
      mydata->paramfirst = this_param;
      mydata->paramlast  = this_param;
      mydata->no_of_params = 1;
   }
   else {
      mydata->paramlast->next = this_param;
      mydata->paramlast = this_param;
      mydata->no_of_params++;
   }

   // Increment number of out parameters if necessary.
   if (isoutput) {
      mydata->no_of_out_params++;
   }

   // Ser shortcuts to make code somewhat less verbose.
   binding    = &(this_param->binding);
   param_info = &(this_param->param_info);
   value      = &(this_param->value);

   // Set up the bindings and parameter information for this parameter.

   // Param_info.datatype, this is a string. We use nameoftype except for
   // UDT for we should use DBTYPE_UDT. And with SQLOLEDB we must use
   // fallbacks for XML and UDT.
   if (this_param->datatype == DBTYPE_UDT) {
       if (mydata->provider == provider_sqlncli) {
          param_info->pwszDataSourceType = SysAllocString(L"DBTYPE_UDT");
       }
       else {
         this_param->datatype = DBTYPE_BYTES;
         param_info->pwszDataSourceType = SysAllocString(L"varbinary");
       }
   }
   else if (this_param->datatype == DBTYPE_XML &&
            mydata->provider != provider_sqlncli) {
      // And different fallback depending on encoding of the XML document.
      xmlcharsettypes charsettype;
      int             charsetpos;

      get_xmlencoding(sv_value, charsettype, charsetpos);

      if (charsettype == eightbit) {
         // If there is an explicit 8-bit encoding, we must use varchar,
         // to avoid "unable to switch the encoding".
         this_param->datatype = DBTYPE_STR;
         param_info->pwszDataSourceType = SysAllocString(L"varchar");
      }
      else {
         this_param->datatype = DBTYPE_WSTR;
         param_info->pwszDataSourceType = SysAllocString(L"nvarchar");

         // Uh-uh, if there is an explicit utf-8 encoding, this will not
         // work out. So...
         if (charsetpos > 0 && charsettype == utf8) {
            // We replace the encoding with ucs-2, because that is what we
            // we actually will send.
            char * str = SvPV_nolen(sv_value);
            str[charsetpos]     = 'u';
            str[charsetpos + 1] = 'c';
            str[charsetpos + 2] = 's';
            str[charsetpos + 3] = '-';
            str[charsetpos + 4] = '2';
         }
     }
   }
   else {
      param_info->pwszDataSourceType = SV_to_BSTR(sv_nameoftype);
   }

   if (paramname && SvOK(paramname)) {
      param_info->pwszName = SV_to_BSTR(paramname);
   }
   else {
      param_info->pwszName = NULL;
   }
   param_info->dwFlags = (isinput  ? DBPARAMFLAGS_ISINPUT : 0) |
                         (isoutput ? DBPARAMFLAGS_ISOUTPUT : 0);
   param_info->bPrecision = 0;
   param_info->bScale     = 0;

   // Binding.
   binding->iOrdinal   = mydata->no_of_params;
   binding->dwMemOwner = DBMEMOWNER_CLIENTOWNED;
   binding->pTypeInfo  = NULL;
   binding->pObject    = NULL;
   binding->pBindExt   = NULL;
   binding->dwFlags    = 0;
   binding->eParamIO   = (isinput  ? DBPARAMIO_INPUT : 0) |
                         (isoutput ? DBPARAMIO_OUTPUT : 0);
   binding->cbMaxLen   = 0;   // For those where it's ignored.
   binding->wType      = this_param->datatype;   // Some will get a BYREF added.
   binding->obLength   = 0;

   // We always bind status and value.
   binding->dwPart    = DBPART_VALUE | DBPART_STATUS;
   binding->obStatus  = mydata->size_param_buffer;
   mydata->size_param_buffer += sizeof(DBSTATUS);
   binding->obValue   = mydata->size_param_buffer;

   switch (this_param->datatype) {
      case DBTYPE_BOOL :
         param_info->ulParamSize = sizeof(BOOL);
         mydata->size_param_buffer += sizeof(long);
         if (! this_param->isnull) {
            value->bit = SvTRUE(sv_value);
         }
         break;

      case DBTYPE_UI1 :
         param_info->ulParamSize = 1;
         mydata->size_param_buffer += 1;
         if (! this_param->isnull) {
            value->tinyint = (BYTE) SvIV(sv_value);
         }
         break;

      case DBTYPE_I2 :
         param_info->ulParamSize = 2;
         mydata->size_param_buffer += 2;
         if (! this_param->isnull) {
            value->smallint = (SHORT) SvIV(sv_value);
         }
         break;

      case DBTYPE_I4 :
         param_info->ulParamSize = 4;
         mydata->size_param_buffer += 4;
         if (! this_param->isnull) {
            value->intval = SvIV(sv_value);
         }
         break;

      case DBTYPE_I8 :
         param_info->ulParamSize = 8;
         mydata->size_param_buffer += 8;
         if (! this_param->isnull) {
            value_OK = SV_to_bigint(sv_value, value->bigint);
         }
         break;

      case DBTYPE_R4 :
         param_info->ulParamSize = 4;
         mydata->size_param_buffer += 4;
         if (! this_param->isnull) {
            value->real = (FLOAT) SvNV(sv_value);
         }
         break;

      case DBTYPE_R8 :
         param_info->ulParamSize = 8;
         mydata->size_param_buffer += 8;
         if (! this_param->isnull) {
            value->floatval = SvNV(sv_value);
         }
         break;

      case DBTYPE_NUMERIC :
         param_info->ulParamSize = sizeof(DB_NUMERIC);
         param_info->bPrecision = precision;
         param_info->bScale     = scale;
         binding->bPrecision = precision;
         binding->bScale     = scale;
         mydata->size_param_buffer += sizeof(DB_NUMERIC);
         if (! this_param->isnull) {
            value_OK = SV_to_decimal(sv_value, precision, scale, value->decimal);
         }
         break;

      case DBTYPE_CY :
         param_info->ulParamSize = sizeof(CY);
         mydata->size_param_buffer += sizeof(CY);
         if (! this_param->isnull) {
            value_OK = SV_to_money(sv_value, value->money);
         }
         break;

      case DBTYPE_DBTIMESTAMP :
         param_info->ulParamSize = sizeof(DBTIMESTAMP);
         mydata->size_param_buffer += sizeof(DBTIMESTAMP);
         if (! this_param->isnull) {
            value_OK = SV_to_datetime(sv_value, value->datetime, olle_ptr);
         }
         if (isoutput) {
            // Set precision, so we can discern datetime/smalldatetime on return.
            if (strcmp(nameoftype, "smalldatetime") == 0) {
               binding->bPrecision = 16;
               binding->bScale     = 0;
            }
            else {
               binding->bPrecision = 23;
               binding->bScale     = 3;
            }
         }
         break;


      case DBTYPE_GUID :
         param_info->ulParamSize = sizeof(GUID);
         mydata->size_param_buffer += sizeof(GUID);
         if (! this_param->isnull) {
            value_OK = SV_to_GUID(sv_value, value->guid);
         }
         break;

      case DBTYPE_STR :
         param_info->ulParamSize = maxlen;
         binding->wType |= DBTYPE_BYREF;
         mydata->size_param_buffer += sizeof(char *);
         binding->dwPart   |= DBPART_LENGTH;
         binding->obLength  = mydata->size_param_buffer;
         mydata->size_param_buffer += sizeof(ULONG);
         if (! this_param->isnull) {
            value_OK = SV_to_char(sv_value, value->varchar,
                                  this_param->value_len);
            this_param->buffer_ptr = value->varchar;
         }
         break;

      case DBTYPE_XML :
         // For XML we may have to add parameter properties.
         add_param_props(olle_ptr, this_param, typeinfo);
         param_info->ulParamSize = maxlen;
         binding->wType |= DBTYPE_BYREF;
         mydata->size_param_buffer += sizeof(char *);
         binding->dwPart   |= DBPART_LENGTH;
         binding->obLength  = mydata->size_param_buffer;
         mydata->size_param_buffer += sizeof(ULONG);
         if (! this_param->isnull) {
            // The return data from SV_to_XML can either an 8-bit or a wide
            // string depending on encoding.
            char * value_ptr;
            BSTR   value_bstr;
            BOOL   is_8bit;
            value_OK = SV_to_XML(sv_value, is_8bit, value_ptr, value_bstr,
                                 this_param->value_len);
            if (is_8bit) {
               value->varchar = value_ptr;
               this_param->buffer_ptr = value->varchar;
            }
            else {
               value->nvarchar = value_bstr;
               this_param->bstr = value->nvarchar;
            }
         }
         break;

      case DBTYPE_WSTR :
         param_info->ulParamSize = maxlen;
         binding->wType |= DBTYPE_BYREF;
         mydata->size_param_buffer += sizeof(WCHAR *);
         binding->dwPart   |= DBPART_LENGTH;
         binding->obLength  = mydata->size_param_buffer;
         mydata->size_param_buffer += sizeof(ULONG);
         if (! this_param->isnull) {
            value->nvarchar = SV_to_BSTR(sv_value);
            this_param->bstr = value->nvarchar;
            this_param->value_len = 2 * SysStringLen(value->nvarchar);
         }
         break;

      case DBTYPE_UDT   :
         // UDT is just like binary, but we have to add parameter properties.
         add_param_props(olle_ptr, this_param, typeinfo);
         // fall-through.
      case DBTYPE_BYTES :
         param_info->ulParamSize = maxlen;
         binding->wType |= DBTYPE_BYREF;
         mydata->size_param_buffer += sizeof(BYTE *);
         binding->dwPart   |= DBPART_LENGTH;
         binding->obLength  = mydata->size_param_buffer;
         mydata->size_param_buffer += sizeof(ULONG);
         if (! this_param->isnull) {
            value_OK = SV_to_binary(sv_value, OptBinaryAsStr(olle_ptr),
                                    istimestamp, value->binary,
                                    this_param->value_len);
            this_param->buffer_ptr = value->binary;
         }
         break;

      case DBTYPE_SQLVARIANT :
         param_info->ulParamSize = sizeof(SSVARIANT);
         mydata->size_param_buffer += sizeof(SSVARIANT);
         if (! this_param->isnull) {
            value_OK = SV_to_ssvariant(sv_value, value->sql_variant,
                                       olle_ptr, this_param->buffer_ptr,
                                       this_param->bstr);
         }
         break;

      case DBTYPE_EMPTY :
         // This value is returned by lookup_type_map, when the type is
         // unknown. This will lead to a warning in the error handler.
         value_OK = FALSE;
         break;

      default :
         // If we come here, this is an internal error in the XS module.
          warn ("Param handling for type %s not implemented yet", nameoftype);
         value_OK = FALSE;
   }
   mydata->all_params_OK &= value_OK;

   if (! value_OK) {
      // There was a conversion error. Issue an error message through the
      // message handler.
      BSTR param_name = param_info->pwszName;
      BSTR data_type = param_info->pwszDataSourceType;
      DBLENGTH stringreplen;
      BSTR stringrep = SV_to_BSTR(sv_value, &stringreplen);
      BSTR msg = SysAllocStringLen(NULL, 200 +
                                   (param_name != NULL ? wcslen(param_name) : 0) +
                                   wcslen(data_type) +
                                   stringreplen);

      if (this_param->datatype != DBTYPE_EMPTY) {
         wsprintf(msg, L"Could not convert Perl value '%s' to type %s for parameter '%s'.",
                   stringrep, data_type, (param_name != NULL ? param_name : L""));
      }
      else {
         wsprintf(msg, L"Data-type name '%s' for parameter '%s' is illegal.",
                  data_type, param_name);
      }

      olledb_message(olle_ptr, -1, 1, 10, msg);
      SysFreeString(stringrep);
      SysFreeString(msg);
      return FALSE;
   }

   return TRUE;
}

//-------------------------------------------------------------------
// set_rowset_properties, this is a subroutine to executebatch.
//-------------------------------------------------------------------
SV  * get_QN_hash(HV * hv,
                  const char * key)
{
   // Retrieves a key value from the QH hash. The value must be defined, and
   // a string value must not be the empty string.
   SV  ** svp;
   SV   * sv = NULL;
   SV   * ret = NULL;

   svp = hv_fetch(hv, key, strlen(key), 0);
   if (svp != NULL) {
       sv = *svp;
   }
   if (sv && SvOK(sv)) {
      if (SvPOK(sv) && SvCUR(sv) >= 1) {
         ret = sv;
      }
      else if (! SvPOK(sv)) {
         ret = sv;
      }
   }

   return ret;
}

void set_rowset_properties (SV           * olle_ptr,
                            internaldata * mydata)
{
    int                  optCommandTimeout = OptCommandTimeout(olle_ptr);
    HV                 * optQN = OptQueryNotification(olle_ptr);
    ICommandProperties * property_ptr;
    DBPROP               property[3];
    int                  no_of_props = 0;
    DBPROPSET            property_set[2];
    int                  no_of_propsets = 0;
    HRESULT              ret;

    if (optCommandTimeout > 0) {
       // There are a lot of properties in DBPROPSET_ROWSET, but we only care
       // about this single one.
       property[0].dwPropertyID = DBPROP_COMMANDTIMEOUT;
       property[0].dwOptions    = DBPROPOPTIONS_REQUIRED;
       property[0].colid        = DB_NULLID;
       VariantInit(&property[0].vValue);
       property[0].vValue.vt    = VT_I4;
       property[0].vValue.lVal  = optCommandTimeout;

       property_set[0].guidPropertySet = DBPROPSET_ROWSET;
       property_set[0].cProperties     = 1;
       property_set[0].rgProperties    = property;

       no_of_propsets++;
    }

    if (optQN) {
       SV   * sv_service;
       SV   * sv_message;
       SV   * sv_timeout;

       no_of_props = 0;

       // First, see if there is a service. Only if there is a service we
       // will submit any query notification at all.
       sv_service = get_QN_hash(optQN, "Service");
       if (sv_service != NULL) {
          if (mydata->provider == provider_sqlncli) {
             property[no_of_props].dwPropertyID = SSPROP_QP_NOTIFICATION_OPTIONS;
             property[no_of_props].dwOptions    = DBPROPOPTIONS_REQUIRED;
             property[no_of_props].colid        = DB_NULLID;
             VariantInit(&property[no_of_props].vValue);
             property[no_of_props].vValue.vt    = VT_BSTR;
             property[no_of_props].vValue.bstrVal  = SV_to_BSTR(sv_service);
             no_of_props++;
          }
          else if (PL_dowarn) {
             BSTR msg = SysAllocString(L"QueryNotification option ignored when provider is SQLOLEDB.");
             olledb_message(olle_ptr, -1, 1, 10, msg);
             SysFreeString(msg);
             sv_service = NULL;
          }
       }
       else if (PL_dowarn && SvTRUE(hv_scalar(optQN))) {
          // If there were other elements in the hash, the user has messed up.
          BSTR msg = SysAllocString(L"The QueryNotification property had elements, but no Service element. No notification was submitted.");
          olledb_message(olle_ptr, -1, 1, 10, msg);
          SysFreeString(msg);
       }

       // We must add a message, so if the user did not provide one, we will.
       sv_message = get_QN_hash(optQN, "Message");
       if (sv_service != NULL) {
          property[no_of_props].dwPropertyID = SSPROP_QP_NOTIFICATION_MSGTEXT;
          property[no_of_props].dwOptions    = DBPROPOPTIONS_REQUIRED;
          property[no_of_props].colid        = DB_NULLID;
          VariantInit(&property[no_of_props].vValue);
          property[no_of_props].vValue.vt    = VT_BSTR;
          property[no_of_props].vValue.bstrVal  =
              (sv_message != NULL ? SV_to_BSTR(sv_message) :
                             SysAllocString(L"Query notification set by Win32::SqlServer"));
          no_of_props++;
       }

       // The timeout on the other hand is optional.
       sv_timeout = get_QN_hash(optQN, "Timeout");
       if (sv_service != NULL && sv_timeout != NULL) {
          property[no_of_props].dwPropertyID = SSPROP_QP_NOTIFICATION_TIMEOUT;
          property[no_of_props].dwOptions    = DBPROPOPTIONS_REQUIRED;
          property[no_of_props].colid        = DB_NULLID;
          VariantInit(&property[no_of_props].vValue);
          property[no_of_props].vValue.vt    = VT_UI4;
          property[no_of_props].vValue.ulVal  = SvIV(sv_timeout);
          no_of_props++;
       }


       // Wipe out the hash.
       hv_clear(optQN);

       if (no_of_props > 0) {
          property_set[no_of_propsets].guidPropertySet = DBPROPSET_SQLSERVERROWSET;
          property_set[no_of_propsets].cProperties     = no_of_props;
          property_set[no_of_propsets].rgProperties    = property;

          no_of_propsets++;
       }
    }

    if (no_of_propsets > 0) {
       // Get a property pointer.
       ret = mydata->cmdtext_ptr->QueryInterface(IID_ICommandProperties,
                                                (void **) &property_ptr);
       check_for_errors(olle_ptr, "cmdtext_ptr->QueryInterface to create Property object", ret);

       ret = property_ptr->SetProperties(no_of_propsets, property_set);
       check_for_errors(NULL, "property_ptr->SetProperties for rowset props", ret);

       property_ptr->Release();
    }

    // We must free up memory allocated to the BSTRs in the QN propset.
    if (optQN) {
       for (int i = 0; i < no_of_props; i++) {
          VariantClear(&property[no_of_props].vValue);
       }
    }
}

//-------------------------------------------------------------------
// $X->executebatch.
//-------------------------------------------------------------------
int executebatch(SV   *olle_ptr,
                 SV   *sv_rows_affected)
{
    internaldata       * mydata = get_internaldata(olle_ptr);
    BOOL                 has_params = (mydata->no_of_params > 0);
    HRESULT              ret;
    paramdata          * current_param;
    DBPARAMBINDINFO    * cur_param_info;
    DBBINDING          * cur_binding;
    ULONG              * param_ordinals;
    ULONG                param_ix = 0;
    ULONG                value_offset;
    ULONG                len_offset;
    ULONG                status_offset;
    BOOL                 final_retval = TRUE;
    ISessionProperties * sess_property_ptr;
    DBPROP               property[1];
    DBPROPSET            property_set[1];
    LONG                 rows_affected;
    DBPARAMS             param_parameter;        // Parameter to cmdtext->Execute.
    SSPARAMPROPS       * ss_param_props = NULL;  // SQL-server specific parameter properies.
    DB_UPARAMS           ss_param_props_cnt = 0;

    // There must be no sesssion_ptr, this indicates that a previous command
    // has not been completely processed.
    if (mydata->cmdtext_ptr != NULL) {
       olle_croak(olle_ptr, "Cannot submit a new batch, when previous batch has not been processed");
    }

    // And check that we have a pending command to execute.
    if (mydata->pending_cmd == NULL) {
       olle_croak(olle_ptr, "There is no pending command to execute. Call initbatch first");
    }

    // And check that we are connect, and connect if auto-connect is set.
    if (mydata->datasrc_ptr == NULL) {
       if (OptAutoConnect(olle_ptr)) {
          if (! do_connect(olle_ptr, TRUE)) {
             return FALSE;
          }
       }
       else {
          olle_croak(olle_ptr, "Not connected to SQL Server, nor is AutoConnect set. Cannot execute batch");
       }
    }

    // If any input parameter failed, to convert, we are not letting you by.
    if (! mydata->all_params_OK) {
        BSTR msg = SysAllocString(
              L"One or more parameters were not convertible. Cannot execute query.");
        olledb_message(olle_ptr, -1, 1, 16, msg);
        SysFreeString(msg);
        free_batch_data(mydata);
        return FALSE;
    }

    // Commands with parameters require a whole lot more of works than
    // those with out.
    if (has_params) {
       // Allocate space for OLE DB's parameter structures and the parameter
       // buffer.
       New(902, mydata->param_info, mydata->no_of_params, DBPARAMBINDINFO);
       New(902, mydata->param_bindings, mydata->no_of_params, DBBINDING);
       New(902, param_ordinals, mydata->no_of_params, ULONG);
       if (mydata->provider == provider_sqlncli) {
          New(902, ss_param_props, mydata->no_of_params, SSPARAMPROPS);
       }

       // Allocate the parameter buffer and initiate it.
       New(902, mydata->param_buffer, mydata->size_param_buffer, BYTE);
       memset(mydata->param_buffer, 0, mydata->size_param_buffer);

       // Iterate over the list to copy the binding and parambindinfo structs,
       // set ordinals and and fill in values to the paramdata buffer.
       current_param  = mydata->paramfirst;
       cur_param_info = mydata->param_info;
       cur_binding    = mydata->param_bindings;
       while (current_param != NULL) {
          // Parameter ordinal.
          param_ordinals[param_ix] = param_ix + 1;

          // Copy structures
          cur_binding[param_ix]    = current_param->binding;
          cur_param_info[param_ix] = current_param->param_info;

          // Get offsets to use.
          value_offset  = current_param->binding.obValue;
          len_offset    = current_param->binding.obLength;
          status_offset = current_param->binding.obStatus;

          // And then fill in the parameter buffer, which is more work.
          if (current_param->isinput) {
             // Write status.
             DBSTATUS * status =
                 (DBSTATUS *) (&mydata->param_buffer[status_offset]);
             * status = (current_param->isnull ? DBSTATUS_S_ISNULL :
                                                 DBSTATUS_S_OK);
             // If not NULL, we need to write input value.
             if (! current_param->isnull) {
                DBLENGTH * len_ptr = NULL;
                if (current_param->binding.dwPart & DBPART_LENGTH) {
                   len_ptr = (DBLENGTH *) (&mydata->param_buffer[len_offset]);
                   * len_ptr = current_param->value_len;
                }

                switch (current_param->datatype) {
                   case DBTYPE_BOOL : {
                      BOOL * buffer_ptr =
                          (BOOL *) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.bit;
                      break;
                   }

                   case DBTYPE_UI1 : {
                      unsigned char * buffer_ptr =
                          (unsigned char *) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.tinyint;
                      break;
                   }

                   case DBTYPE_I2 : {
                      short * buffer_ptr =
                          (short *) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.smallint;
                      break;
                   }

                   case DBTYPE_I4 : {
                      long * buffer_ptr =
                          (long *) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.intval;
                      break;
                   }

                   case DBTYPE_I8 : {
                      LONGLONG * buffer_ptr =
                          (LONGLONG *) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.bigint;
                      break;
                   }

                   case DBTYPE_R4 : {
                      float * buffer_ptr =
                         (float *) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.real;
                      break;
                   }

                   case DBTYPE_R8 : {
                      double * buffer_ptr =
                         (double *) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.floatval;
                      break;
                   }

                   case DBTYPE_NUMERIC : {
                      DB_NUMERIC  * buffer_ptr =
                          (DB_NUMERIC *) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.decimal;
                      break;
                   }

                   case DBTYPE_CY : {
                      CY * buffer_ptr = (CY *) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.money;
                      break;
                   }

                   case DBTYPE_DBTIMESTAMP : {
                      DBTIMESTAMP * buffer_ptr =
                         (DBTIMESTAMP *) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.datetime;
                      break;
                   }

                   case DBTYPE_GUID : {
                      GUID * buffer_ptr =
                          (GUID *) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.guid;
                      break;
                   }

                   case DBTYPE_STR : {
                      char ** buffer_ptr =
                           (char **) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.varchar;
                      break;
                   }

                   case DBTYPE_WSTR : {
                      BSTR * buffer_ptr =
                             (BSTR *) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.nvarchar;
                      break;
                   }

                   case DBTYPE_XML :
                   // XML may have either of a varchar or an nvarchar pointer.
                   // We can tell from our saved buffer_ptr:
                   if (current_param->buffer_ptr != NULL) {
                      char ** buffer_ptr =
                           (char **) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.varchar;
                   }
                   else {
                      BSTR * buffer_ptr =
                             (BSTR *) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.nvarchar;
                   }
                   break;

                   case DBTYPE_UDT   :
                   case DBTYPE_BYTES : {
                      BYTE ** buffer_ptr =
                           (BYTE **) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.binary;
                      break;
                   }

                   case DBTYPE_SQLVARIANT : {
                      SSVARIANT * buffer_ptr =
                         (SSVARIANT *) (&mydata->param_buffer[value_offset]);
                      * buffer_ptr = current_param->value.sql_variant;
                      break;
                   }

                   default :
                     olle_croak(olle_ptr, "Internal error: unhandled type %d",
                                current_param->datatype);
                     break;
                }
             }
             /* Good debug,
             wprintf(L"Param_name = %s, status = %d, value = %d.\n",
                  current_param->param_info.pwszName,
                  current_param->binding.obStatus,
                  current_param->binding.obValue);
             */
          }

          // And finally, fill in parameter properties for XML and UDT.
          if (current_param->param_props_cnt > 0) {
              DBPROPSET  * propset;
              New(902, propset, current_param->param_props_cnt, DBPROPSET);
              propset->rgProperties = current_param->param_props;
              propset->cProperties = current_param->param_props_cnt;
              propset->guidPropertySet = DBPROPSET_SQLSERVERPARAMETER;
              ss_param_props[ss_param_props_cnt].rgPropertySets = propset;
              ss_param_props[ss_param_props_cnt].cPropertySets = 1;
              ss_param_props[ss_param_props_cnt].iOrdinal =
                  param_ordinals[param_ix];
              ss_param_props_cnt++;
          }

          // Move to next.
          current_param = current_param->next;
          param_ix++;
       }

       // Must allocate space for bindstatus.
       New(902, mydata->param_bind_status, mydata->no_of_params, DBBINDSTATUS);
    }   // if has_params

    // We need a session object.
    ret = mydata->datasrc_ptr->CreateSession(NULL, IID_IDBCreateCommand,
                                         (IUnknown **) &(mydata->session_ptr));
    check_for_errors(olle_ptr, "datasrc_ptr->CreateSession for session object", ret);

    // We need a property object for the session
    ret = mydata->session_ptr->QueryInterface(IID_ISessionProperties,
                                             (void **) &sess_property_ptr);
    check_for_errors(olle_ptr, "session_ptr->QueryInterface to create Property object", ret);

    // We always want the SQL Server-native representation of variant data.
    property[0].dwPropertyID   = SSPROP_ALLOWNATIVEVARIANT;
    property[0].dwOptions = DBPROPOPTIONS_REQUIRED;
    property[0].colid     = DB_NULLID;
    VariantInit(&property[0].vValue);
    property[0].vValue.vt      = VT_BOOL;
    property[0].vValue.boolVal = VARIANT_TRUE;

    property_set[0].guidPropertySet = DBPROPSET_SQLSERVERSESSION;
    property_set[0].cProperties     = 1;
    property_set[0].rgProperties    = property;

    ret = sess_property_ptr->SetProperties(1, property_set);
    check_for_errors(NULL, "property_ptr->SetProperties for ssvariant prop", ret);

    sess_property_ptr->Release();

    // Command-text interface.
    ret = mydata->session_ptr->CreateCommand(NULL, IID_ICommandText,
                                         (IUnknown **)  &(mydata->cmdtext_ptr));
    check_for_errors(olle_ptr, "session_ptr->CreateCommand for command-text object", ret);

    // Set rowset properties from Win32::SqlServer options.
    set_rowset_properties(olle_ptr, mydata);

    // Set the command text.
    ret = mydata->cmdtext_ptr->SetCommandText(DBGUID_SQL, mydata->pending_cmd);
    check_for_errors(olle_ptr, "cmdtext_ptr->SetCommandText", ret);

    // Again, extra stuff for commands with parameters
    if (has_params) {
       // Command-with-parameter interface
       ret = mydata->cmdtext_ptr->QueryInterface(IID_ICommandWithParameters,
                                             (void **) &(mydata->paramcmd_ptr));
       check_for_errors(olle_ptr, "cmdtext_ptr->QueryInterface for ICommandWithParameters", ret);

       // Set parameter info. Here we permit execution to proceed in case of
       // errors, as it could be user errors like using the xml datatype with
       // SQLEOLEDB.
       ret = mydata->paramcmd_ptr->SetParameterInfo(mydata->no_of_params,
                                                    param_ordinals,
                                                    mydata->param_info);
       check_for_errors(olle_ptr, "paramcmd_ptr->SetParameterInfo", ret,
                        FALSE);

       if (SUCCEEDED(ret) && ss_param_props_cnt > 0) {
          ret = mydata->cmdtext_ptr->QueryInterface(IID_ISSCommandWithParameters,
                                       (void **) &(mydata->ss_paramcmd_ptr));
          check_for_errors(olle_ptr, "paramcmd_ptr->QueryInterface for ISSCommandWithParameters", ret);

          ret = mydata->ss_paramcmd_ptr->SetParameterProperties(
                                  ss_param_props_cnt, ss_param_props);
          check_for_errors(olle_ptr, "ss_paramcmd_ptr->SetParameterProperties", ret);
       }

       if (SUCCEEDED(ret)) {
          // Get accessor interface.
          ret = mydata->paramcmd_ptr->QueryInterface(IID_IAccessor,
                                             (void **) &(mydata->paramaccess_ptr));
          check_for_errors(olle_ptr, "paramcmd->QueryInterace for IAccessor", ret);

          // And get the accessor itself.
          ret = mydata->paramaccess_ptr->CreateAccessor(
                DBACCESSOR_PARAMETERDATA, mydata->no_of_params,
                mydata->param_bindings, mydata->size_param_buffer,
                &(mydata->param_accessor), mydata->param_bind_status);
          check_for_errors(olle_ptr, "paramacces_ptr->CreateAccessor", ret);

          param_parameter.pData = mydata->param_buffer;
          param_parameter.cParamSets = 1;
          param_parameter.hAccessor = mydata->param_accessor;
       }
    }

    if (SUCCEEDED(ret)) {
       // Now execute the command. Again, proceed on all errors, so we get by
       // the famous "multi-step errors".
       ret = mydata->cmdtext_ptr->Execute(NULL, IID_IMultipleResults,
                                          (has_params ? &param_parameter : NULL),
                                          &rows_affected,
                                          (IUnknown **) &(mydata->results_ptr));
       check_for_errors(olle_ptr, "cmdtext_ptr->Execute", ret, FALSE);
    }

    // check_for_errors returns if the call fails, because one or
    // more parameter could not convert. We should not croak on this,
    // but we do cancel the batch.
    if (FAILED(ret)) {
       final_retval = FALSE;
       free_batch_data(mydata);
    }

    // Return rows_affected if required.
    if (sv_rows_affected != NULL) {
        sv_setiv(sv_rows_affected, rows_affected);
    }

    // Some cleaning up.
    if (has_params) {
       Safefree(param_ordinals);
    }

    if (ss_param_props != NULL) {
       for (DB_UPARAMS ix = 0; ix < ss_param_props_cnt; ix++) {
          Safefree(ss_param_props[ix].rgPropertySets);
       }
       Safefree(ss_param_props);
    }

    return final_retval;
}

//====================================================================
// Get data from SQL Server. First $X->nextresultset, then conversion
// routines to from SQL Server types to SV, then extract data, a common
// helper to $X->nextrow and $X->outputparamerers.
//====================================================================
int nextresultset (SV * olle_ptr,
                   SV * sv_rows_affected)
{
    internaldata * mydata = get_internaldata(olle_ptr);
    LONG           rows_affected;
    int            more_results;
    HRESULT        ret;
    IColumnsInfo*  columns_info_ptr  = NULL;
    ULONG          no_of_cols;
    ULONG          bind_offset = 0;

    // There must not a cmttext_ptr, else there is no command being processed.
    if (mydata->cmdtext_ptr == NULL) {
       olle_croak(olle_ptr, "Cannot call nextresultset without an active command. Call executebatch first");
    }

    // If there are unfetched rows in the previous result set, this is an
    // error.
    if (mydata->rowset_ptr != NULL) {
       olle_croak(olle_ptr, "Cannot call nextresultset with unfetched rows. Call nextrow to get all rows, or call cancelresulset");
    }

    // We are not guaranteed to have a results pointer, but this condition
    // means that there are no more results.
    if (mydata->results_ptr == NULL) {
       more_results = FALSE;
    }
    else {
       // Get next result set. The assumption is here that if we get an
       // SQL error we should continue and give caller what we have. Other
       // errors should lead to a immiedate stop, and we assume that we
       // sooner or later get a DB_S_NORESULT.
       ret = mydata->results_ptr->GetResult(NULL, 0, IID_IRowset, &rows_affected,
                                          (IUnknown **) &(mydata->rowset_ptr));
       check_for_errors(olle_ptr, "results_ptr->GetResults", ret);
       more_results = (ret != DB_S_NORESULT);
    }

    // Do we now have an active result set?
    mydata->have_resultset = more_results;

    // A result set usually comes with a rowset, but it is just a count or
    // a message, it does not.
    if (more_results && mydata->rowset_ptr != NULL) {
       // Get ColumnsInfo interface.
       ret = mydata->rowset_ptr->QueryInterface(IID_IColumnsInfo,
                                                (void **) &columns_info_ptr);
       check_for_errors(olle_ptr,
                        "rowset_ptr->QueryInterface for column info", ret);

       // Get columninfo buffer.
       ret = columns_info_ptr->GetColumnInfo(&no_of_cols,
                                             &(mydata->column_info),
                                             &(mydata->colname_buffer));
       check_for_errors(olle_ptr, "columns_info_ptr->GetColumnInfo", ret);
       mydata->no_of_cols = no_of_cols;

       // Don't need this interface any more.
       columns_info_ptr->Release();

       // We need the col_binding array.
       New(902, mydata->col_bindings, no_of_cols, DBBINDING);

       // Iterate over the columns to set up the bindings.
       for (ULONG j = 0; j < no_of_cols; j++) {
          // These fields are the same for all data types.
          mydata->col_bindings[j].iOrdinal  = j+1;
          mydata->col_bindings[j].dwMemOwner = DBMEMOWNER_CLIENTOWNED;
          mydata->col_bindings[j].pTypeInfo = NULL;
          mydata->col_bindings[j].pObject   = NULL;
          mydata->col_bindings[j].pBindExt  = NULL;
          mydata->col_bindings[j].dwFlags   = 0;
          mydata->col_bindings[j].eParamIO  = DBPARAMIO_NOTPARAM;
          mydata->col_bindings[j].cbMaxLen  = 0;   // For those where it ignoreed.
          mydata->col_bindings[j].wType     = mydata->column_info[j].wType;  // BYREF may be added later.

          // We always bind status and value.
          mydata->col_bindings[j].dwPart    = DBPART_VALUE | DBPART_STATUS;
          mydata->col_bindings[j].obStatus  = bind_offset;
          bind_offset += sizeof(DBSTATUS);
          mydata->col_bindings[j].obValue   = bind_offset;
          // The rest depends on the data type.
          switch (mydata->column_info[j].wType) {
             case DBTYPE_BOOL :
                bind_offset += sizeof(BOOL);
                break;

             case DBTYPE_UI1 :
                bind_offset += 1;
                break;

             case DBTYPE_I2 :
                bind_offset += 2;
                break;

             case DBTYPE_I4 :
                bind_offset += 4;
                break;

             case DBTYPE_R4 :
                bind_offset += 4;
                break;

             case DBTYPE_R8 :
                bind_offset += 8;
                break;

             case DBTYPE_I8 :
                bind_offset += 8;
                break;

             case DBTYPE_CY :
                bind_offset += sizeof(CY);
                break;

             case DBTYPE_NUMERIC :
                mydata->col_bindings[j].bPrecision =
                    mydata->column_info[j].bPrecision;
                mydata->col_bindings[j].bScale     =
                    mydata->column_info[j].bScale;
                bind_offset += sizeof(DB_NUMERIC);
                break;

             case DBTYPE_GUID :
                bind_offset += sizeof(GUID);
                break;

             case DBTYPE_DBTIMESTAMP :
                mydata->col_bindings[j].bPrecision =
                    mydata->column_info[j].bPrecision;
                mydata->col_bindings[j].bScale     =
                    mydata->column_info[j].bScale;
                bind_offset += sizeof(DBTIMESTAMP);
                break;

             case DBTYPE_UDT   :
             case DBTYPE_BYTES :
                mydata->col_bindings[j].wType    |= DBTYPE_BYREF;
                bind_offset += sizeof(BYTE *);
                mydata->col_bindings[j].dwPart   |= DBPART_LENGTH;
                mydata->col_bindings[j].obLength  = bind_offset;
                bind_offset += sizeof(ULONG);
                break;

             case DBTYPE_STR :
                mydata->col_bindings[j].wType    |= DBTYPE_BYREF;
                bind_offset += sizeof(char *);
                mydata->col_bindings[j].dwPart   |= DBPART_LENGTH;
                mydata->col_bindings[j].obLength  = bind_offset;
                bind_offset += sizeof(ULONG);
                break;

             case DBTYPE_SQLVARIANT :
                bind_offset += sizeof(SSVARIANT);
                break;

             case DBTYPE_XML :
                mydata->col_bindings[j].wType    |= DBTYPE_BYREF;
                bind_offset += sizeof(WCHAR *);
                mydata->col_bindings[j].dwPart   |= DBPART_LENGTH;
                mydata->col_bindings[j].obLength  = bind_offset;
                bind_offset += sizeof(ULONG);
                break;

             case DBTYPE_WSTR :
             default          :
                if (mydata->column_info[j].wType != DBTYPE_WSTR) {
                   warn("Warning: Unexpected datatype %d, handled as nvarchar.",
                        mydata->column_info[j].wType);
                   mydata->col_bindings[j].wType = DBTYPE_WSTR;
                   mydata->column_info[j].wType = DBTYPE_WSTR;
                }
                mydata->col_bindings[j].wType     |= DBTYPE_BYREF;
                bind_offset += sizeof(WCHAR *);
                mydata->col_bindings[j].dwPart    |= DBPART_LENGTH;
                mydata->col_bindings[j].obLength   = bind_offset;
                bind_offset += sizeof(ULONG);
                break;
          }
       }

       // Save the final offset and allocate space for data buffer.
       mydata->size_data_buffer = bind_offset;
       New(902, mydata->data_buffer, bind_offset, BYTE);

       // Get the accessor interface.
       ret = mydata->rowset_ptr->QueryInterface(IID_IAccessor,
                                    (void **) &(mydata->rowaccess_ptr));
       check_for_errors(olle_ptr,
                        "rowset_ptr->QueryInterface for row accessor", ret);

       // Must allocate space for DBBINDSTATUS.
       New(902, mydata->col_bind_status, no_of_cols, DBBINDSTATUS);

       ret = mydata->rowaccess_ptr->CreateAccessor(DBACCESSOR_ROWDATA,
                                                   no_of_cols,
                                                   mydata->col_bindings, 0,
                                                   &(mydata->row_accessor),
                                                   mydata->col_bind_status);
       check_for_errors(olle_ptr, "rowaccess_ptr->CreateAccessor", ret);
    }
    else if (! more_results) {
       if (mydata->no_of_out_params == 0) {
          // If there are no output parameters, we can free resources bound
          // the by current batch here and now.
          free_batch_data(mydata);
       }
       else {
         // Else just make output parameters available.
         mydata->params_available = TRUE;
       }
    }

    // Return rows_affected if required.
    if (sv_rows_affected != NULL) {
        sv_setiv(sv_rows_affected, rows_affected);
    }

    return more_results;
}

//---------------------------------------------------------------------
// Conversion-to-SV routines. These routines convert a non-trivial value
// from SQL Server a suitable SV. In most cases there is an option that
// determines the datatype/format for the Perl value.
//---------------------------------------------------------------------
SV * bigint_to_SV (LONGLONG       bigintval,
                   formatoptions  opts)
{
   if (opts.DecimalAsStr) {
      char str[20];
      sprintf(str, "%I64d", bigintval);
      return newSVpv(str, 0);
   }
   else {
      return newSVnv((double) bigintval);
   }
}

SV * binary_to_SV (BYTE         * binaryval,
                   DBLENGTH      len,
                   formatoptions opts)
{
   SV   * perl_value;

   if (opts.BinaryAsStr != bin_binary) {
       DBLENGTH         strlen;
       char           * strsans0x;
       char           * str0x;
       DBSTATUS         strstatus;
       HRESULT          ret;

       New(902, strsans0x, 2 * len + 1, char);

       ret = data_convert_ptr->DataConvert(
             DBTYPE_BYTES, DBTYPE_STR, len, &strlen,
             binaryval, strsans0x, 2 * len + 1, DBSTATUS_S_OK, &strstatus,
             NULL, NULL, 0);
       check_convert_errors("Convert binary-to-str", strstatus, ret);

       if (opts.BinaryAsStr == bin_string0x) {
          New(902, str0x, strlen + 3, char);
          sprintf(str0x, "0x%s", strsans0x);
          perl_value = newSVpvn(str0x, strlen + 2);
          Safefree(str0x);
       }
       else {
          perl_value = newSVpvn(strsans0x, strlen);
       }
       Safefree(strsans0x);
   }
   else {
       perl_value = newSVpvn((char *) binaryval, len);
   }

   return perl_value;
}

SV * bit_to_SV (VARIANT_BOOL bitval)
{
   return newSViv(bitval == 0 ? 0 : 1);
}


SV * datetime_to_SV (SV          * olle_ptr,
                     DBTIMESTAMP   datetime,
                     formatoptions opts,
                     BYTE          precision,
                     BYTE          scale)
{

   SV         * perl_value;

   // For dates there is a multitude of options.
   switch (opts.DatetimeOption) {
      case dto_hash : {
            HV * hv = newHV();
            SV * year = newSViv(datetime.year);
            SV * month = newSViv(datetime.month);
            SV * day = newSViv(datetime.day);
            SV * hour = newSViv(datetime.hour);
            SV * minute = newSViv(datetime.minute);
            SV * second = newSViv(datetime.second);
            SV * fraction = newSViv(datetime.fraction / 1000000);

            hv_store(hv, "Year",     strlen("Year"),     year, 0);
            hv_store(hv, "Month",    strlen("Month"),    month, 0);
            hv_store(hv, "Day",      strlen("Day"),      day, 0);
            hv_store(hv, "Hour",     strlen("Hour"),     hour, 0);
            hv_store(hv, "Minute",   strlen("Minute"),   minute, 0);
            hv_store(hv, "Second",   strlen("Second"),   second, 0);
            hv_store(hv, "Fraction", strlen("Fraction"), fraction, 0);

            perl_value = newSV(NULL);
            sv_setsv(perl_value, sv_2mortal(newRV_noinc((SV *) hv)));
         }
         break;

      case dto_iso : {
            DBLENGTH       strlen;
            char           str[30];
            DBSTATUS       strstatus;
            HRESULT        ret;

            ret = data_convert_ptr->DataConvert(
                  DBTYPE_DBTIMESTAMP, DBTYPE_STR, sizeof(DBTIMESTAMP),
                  &strlen, &datetime, &str, 30, DBSTATUS_S_OK, &strstatus,
                  precision, scale, 0);
            check_convert_errors("Convert datetime-to-str", strstatus, ret);

            // For datetime DataConvert does not fill in msecs if they are zero.
            if (precision == 23 && strlen < 23) {
               sprintf(&str[19], ".000");
            }
            perl_value = newSVpvn(str, precision);
         }
         break;

      case dto_regional : {
            // This conversion requires a double conversion. First to DATE.
            // and then to string.
            DATE           dateval;
            BSTR           bstr;
            DBSTATUS       dbstatus;
            HRESULT        ret;

            ret = data_convert_ptr->DataConvert(
                  DBTYPE_DBTIMESTAMP, DBTYPE_DATE, sizeof(DBTIMESTAMP),
                  NULL, &datetime, &dateval, sizeof(DATE), DBSTATUS_S_OK,
                  &dbstatus, precision, scale, 0);
            check_convert_errors("Convert datetime-to-date", dbstatus, ret);

            ret = VarBstrFromDate(dateval, 0, 0, &bstr);
            check_convert_errors("Convert date-to-str", dbstatus, ret);

            perl_value = BSTR_to_SV(bstr);
            SysFreeString(bstr);
         }
         break;

      case dto_float : {
            DATE           dateval;
            DBSTATUS       dbstatus;
            HRESULT        ret;

            ret = data_convert_ptr->DataConvert(
                  DBTYPE_DBTIMESTAMP, DBTYPE_DATE, sizeof(DBTIMESTAMP),
                  NULL, &datetime, &dateval, sizeof(DATE), DBSTATUS_S_OK,
                  &dbstatus, NULL, NULL, 0);
            check_convert_errors("Convert datetime-to-date", dbstatus, ret);

            perl_value = newSVnv(dateval);
        }
        break;

      case dto_strfmt : {
            struct tm tm_date;
            size_t     len;
            size_t     msec_len = 0;
            char       str[256];

            if (opts.DateFormat == NULL || ! *opts.DateFormat) {
               olle_croak(olle_ptr, "Datetime option set to dt_strfmt, but there is no format defined");
            }

            // Move over data to the tm_struct.
            tm_date.tm_hour  = datetime.hour;
            tm_date.tm_isdst = 0; // Seriously, we don't know.
            tm_date.tm_mday  = datetime.day;
            tm_date.tm_min   = datetime.minute;
            tm_date.tm_mon   = datetime.month - 1;
            tm_date.tm_sec   = datetime.second;
            tm_date.tm_wday  = 0;
            tm_date.tm_yday  = 0;
            tm_date.tm_year  = datetime.year - 1900;

            // Convert the beast
            len = strftime(str, 256, opts.DateFormat, &tm_date);
            if (len <= 0) {
               olle_croak(olle_ptr, "strftime failed for dateFormat '%s'", opts.DateFormat);
            }

            // Are we also requested to format milliseconds?
            if (scale > 0 && opts.MsecFormat && * opts.MsecFormat) {
               msec_len = _snprintf(&str[len], 50 - len, opts.MsecFormat,
                               datetime.fraction / 1000000);
               if (msec_len <= 0) {
                  olle_croak(olle_ptr, "_snprintf failed for msecFormat '%s'", opts.MsecFormat);
               }
            }

            perl_value = newSVpv(str, len + msec_len);
         }
         break;

      default :
         olle_croak(olle_ptr, "Illegal value for DatetimeOption %d", opts.DatetimeOption);

   }

   return perl_value;
}

SV * decimal_to_SV (DB_NUMERIC    decimalval,
                    formatoptions opts)
{
   DBLENGTH       sstrlen;
   char           str[50];
   DBSTATUS       status;
   HRESULT        ret;

   if (opts.DecimalAsStr) {
      ret = data_convert_ptr->DataConvert(
            DBTYPE_NUMERIC, DBTYPE_STR, sizeof(DB_NUMERIC), &sstrlen,
            &decimalval, &str, 50, DBSTATUS_S_OK, &status, NULL, NULL, 0);
      check_convert_errors("Convert decimal-to-str", status, ret);

      return newSVpvn(str, sstrlen);
   }
   else {
      double dbl;

      ret = data_convert_ptr->DataConvert(
           DBTYPE_NUMERIC, DBTYPE_R8, sizeof(DB_NUMERIC), NULL,
           &decimalval, &dbl, NULL, DBSTATUS_S_OK, &status, NULL, NULL, 0);
      check_convert_errors("Convert decimal-to-float", status, ret);

      return newSVnv(dbl);
   }
}


SV * GUID_to_SV (GUID    guid)
{
    DBLENGTH  strlen;
    char      str[40];
    DBSTATUS  strstatus;
    HRESULT   ret;

    ret = data_convert_ptr->DataConvert(DBTYPE_GUID, DBTYPE_STR, sizeof(GUID),
                                        &strlen, &guid, &str, 40,
                                        DBSTATUS_S_OK, &strstatus,
                                        NULL, NULL, 0);
    check_convert_errors ("Convert GUID to STR", strstatus, ret);

    return newSVpvn(str, strlen);
}


SV * money_to_SV (CY            moneyval,
                  formatoptions opts)
{
    DBLENGTH       sstrlen;
    char           str[50];
    DBSTATUS       status;
    HRESULT        ret;

    if (opts.DecimalAsStr) {
       ret = data_convert_ptr->DataConvert(
             DBTYPE_CY, DBTYPE_STR, sizeof(CY), &sstrlen,
             &moneyval, &str, 50, DBSTATUS_S_OK, &status, NULL, NULL, 0);
       check_convert_errors("Convert money-to-str", status, ret);

       return newSVpvn(str, sstrlen);
    }
    else {
       double dbl;

       ret = data_convert_ptr->DataConvert(
            DBTYPE_CY, DBTYPE_R8, sizeof(CY), NULL,
            &moneyval, &dbl, NULL, DBSTATUS_S_OK, &status, NULL, NULL, 0);
       check_convert_errors("Convert money-to-float", status, ret);

       return newSVnv(dbl);
    }
}


SV * ssvariant_to_SV(SV          * olle_ptr,
                     SSVARIANT     ssvar,
                     formatoptions opts)
{
   SV * perl_value;

   switch (ssvar.vt) {
      case VT_SS_EMPTY :
      case VT_SS_NULL  :
         perl_value = newSVsv(&PL_sv_undef);
         break;

      case VT_SS_UI1 :
         perl_value = newSViv(ssvar.bTinyIntVal);
         break;

      case VT_SS_I2 :
         perl_value = newSViv(ssvar.sShortIntVal);
         break;

      case VT_SS_I4 :
         perl_value = newSViv(ssvar.lIntVal);
         break;

      case VT_SS_I8 :
         perl_value = bigint_to_SV(ssvar.llBigIntVal, opts);
         break;

      case VT_SS_R4 :
         perl_value = newSVnv(ssvar.fltRealVal);
         break;

      case VT_SS_R8 :
         perl_value = newSVnv(ssvar.dblFloatVal);
         break;

      case VT_SS_MONEY :
      case VT_SS_SMALLMONEY :
         perl_value = money_to_SV(ssvar.cyMoneyVal, opts);
         break;

       case VT_SS_WSTRING    :
       case VT_SS_WVARSTRING :
          perl_value = BSTR_to_SV(ssvar.NCharVal.pwchNCharVal,
                                  ssvar.NCharVal.sActualLength / 2);
          OLE_malloc_ptr->Free(ssvar.NCharVal.pwchNCharVal);
          if (ssvar.NCharVal.pwchReserved != NULL) {
             OLE_malloc_ptr->Free(ssvar.NCharVal.pwchReserved);
          }
          break;

       case VT_SS_STRING    :
       case VT_SS_VARSTRING :
          perl_value = newSVpv(ssvar.CharVal.pchCharVal,
                               ssvar.CharVal.sActualLength);
          OLE_malloc_ptr->Free(ssvar.CharVal.pchCharVal);
          if (ssvar.NCharVal.pwchReserved != NULL) {
             OLE_malloc_ptr->Free(ssvar.NCharVal.pwchReserved);
          }
          break;

       case VT_SS_BIT  :
          perl_value = bit_to_SV(ssvar.fBitVal);
          break;

       case VT_SS_GUID : {
          GUID guid;
          memcpy(&guid, ssvar.rgbGuidVal, 16);
          perl_value = GUID_to_SV(guid);
          break;
       }

       case VT_SS_NUMERIC :
       case VT_SS_DECIMAL :
          perl_value = decimal_to_SV(ssvar.numNumericVal, opts);
          break;

       case VT_SS_DATETIME      :
          perl_value = datetime_to_SV(olle_ptr, ssvar.tsDateTimeVal,
                                      opts, 23, 3);
          break;

       case VT_SS_SMALLDATETIME :
          perl_value = datetime_to_SV(olle_ptr, ssvar.tsDateTimeVal,
                                      opts, 16, 0);
          break;

       case VT_SS_BINARY    :
       case VT_SS_VARBINARY :
          perl_value = binary_to_SV(ssvar.BinaryVal.prgbBinaryVal,
                                    ssvar.BinaryVal.sActualLength, opts);
          OLE_malloc_ptr->Free(ssvar.BinaryVal.prgbBinaryVal);
          break;

       default : {
          char str[40];
          sprintf(str, "Unsupported value for SSVARIANT.vt: %d", ssvar.vt);
          perl_value = newSVpv(str, 0);
          break;
       }
   }

   return perl_value;
}


//-----------------------------------------------------------------------
// This routine converts a column/parameter value from the the data buffer
// with help of the binding information into an SV. This one is called both
// from nextrow and getoutputparameters.
//------------------------------------------------------------------------
void extract_data(SV           * olle_ptr,
                  formatoptions  opts,
                  BOOL           is_param,
                  BSTR           valuename,
                  DBTYPE         datatype,
                  DBBINDING      binding,
                  DBBINDSTATUS   bind_status,
                  BYTE         * data_buffer,
                  SV           * &perl_value)
{

    DBSTATUS       value_status = *((DBSTATUS *) &data_buffer[binding.obStatus]);
    DBBYTEOFFSET   value_offset = binding.obValue;
    ULONG          value_len = 0;

    if (binding.dwPart & DBPART_LENGTH) {
       value_len = * ((ULONG *) &data_buffer[binding.obLength]);
    }

   /* {
                    char * str =  * ((char **) &data_buffer[value_offset]);
                    warn("datastr_ptr = %x '%s'\n", str, valuename);
    } */

    switch (value_status) {
       case DBSTATUS_S_ISNULL :
            perl_value = newSVsv(&PL_sv_undef);
            break;

       case DBSTATUS_S_TRUNCATED : {
            char * tmp = BSTR_to_char(valuename);
            warn("Value of column/parameter '%s' was truncated.", tmp);
            Safefree(tmp);
       }
            // fall-through.
       case DBSTATUS_S_OK :
          switch (datatype) {
             case DBTYPE_BOOL      : {
                BOOL value = * ((BOOL *) &data_buffer[value_offset]);
                perl_value = bit_to_SV(value);
                break;
             }

             case DBTYPE_UI1       : {
                unsigned char value =
                     * ((unsigned char *) &data_buffer[value_offset]);
                perl_value = newSViv(value);
                break;
             }

             case DBTYPE_I2        : {
                short value = * ((short *) &data_buffer[value_offset]);
                perl_value = newSViv(value);
                break;
             }

             case DBTYPE_I4        :  {
                long value = * ((long *) &data_buffer[value_offset]);
                perl_value = newSViv(value);
                break;
             }

             case DBTYPE_R4       : {
                float value = * ((float *) &data_buffer[value_offset]);
                perl_value = newSVnv(value);
                break;
             }

             case DBTYPE_R8       : {
                double value = * ((double *) &data_buffer[value_offset]);
                perl_value = newSVnv(value);
                break;
             }

             case DBTYPE_I8       : {
                LONGLONG value = * ((LONGLONG *) &data_buffer[value_offset]);
                perl_value = bigint_to_SV(value, opts);
                break;
             }

             case DBTYPE_CY       : {
                CY value = * ((CY *) &data_buffer[value_offset]);
                perl_value = money_to_SV(value, opts);
                break;
             }

             case DBTYPE_NUMERIC  : {
                DB_NUMERIC value = * ((DB_NUMERIC *) &data_buffer[value_offset]);
                perl_value = decimal_to_SV(value, opts);
                break;
             }

             case DBTYPE_DBTIMESTAMP : {
                DBTIMESTAMP value = * ((DBTIMESTAMP *) &data_buffer[value_offset]);
                perl_value = datetime_to_SV(olle_ptr, value, opts,
                                            binding.bPrecision, binding.bScale);
                break;
             }

             case DBTYPE_GUID : {
                GUID value = * ((GUID *) &data_buffer[value_offset]);
                perl_value = GUID_to_SV(value);
                break;
             }

             case DBTYPE_SQLVARIANT  : {
                SSVARIANT ssvar = * ((SSVARIANT *) &data_buffer[value_offset]);
                perl_value = ssvariant_to_SV(olle_ptr, ssvar, opts);
                break;
             }

             case DBTYPE_UDT   :
             case DBTYPE_BYTES : {
                BYTE ** byteptr =  ((BYTE **) &data_buffer[value_offset]);
                perl_value = binary_to_SV(* byteptr, value_len, opts);
                OLE_malloc_ptr->Free(* byteptr);
                * byteptr = NULL; // Clear entry in buffer, since ptr no longer valid.
                break;
             }

             case DBTYPE_STR   : {
                char ** strptr =  ((char **) &data_buffer[value_offset]);
                perl_value = newSVpvn(* strptr, value_len);
                OLE_malloc_ptr->Free(* strptr);
                * strptr = NULL;
                break;
             }

             case DBTYPE_XML   : {
                // For XML there is BOM, that we should ignore.
                WCHAR ** strptr =  ((WCHAR **) &data_buffer[value_offset]);
                WCHAR * xmlptr =  * strptr;
                perl_value = BSTR_to_SV(xmlptr + 1, value_len / 2 - 1);
                OLE_malloc_ptr->Free(* strptr);
                * strptr = NULL;
                break;
             }

             case DBTYPE_WSTR  : {
                WCHAR ** strptr =  ((WCHAR **) &data_buffer[value_offset]);
                perl_value = BSTR_to_SV(* strptr, value_len / 2);
                OLE_malloc_ptr->Free(* strptr);
                * strptr = NULL;
                break;
             }

             default :
                olle_croak(olle_ptr, "Internal error: Unexpected data type %d in extract_data", datatype);
                break;
          }
          break;

       case DBSTATUS_E_UNAVAILABLE :
       // This may happen with a parameter value, if the command fails,
       // in which case we just set undef. This "should not happen" with a
       // column, so for a column we should croak on this. Whence the
       // funky placement of break for a half fall-through.
          if (is_param) {
              perl_value = newSVsv(&PL_sv_undef);
              break;
          }

       default : {
          char  msg[200];
          char  * tmp = BSTR_to_char(valuename);
          sprintf(msg, "Extraction of param/col '%s'", tmp);
          Safefree(tmp);
          check_convert_errors(msg, value_status, bind_status, S_OK);
       }
   }
}

//--------------------------------------------------------------------
// $X->nextrow
//--------------------------------------------------------------------
int nextrow (SV   * olle_ptr,
             SV   * hashref,
             SV   * arrayref)
{
    internaldata * mydata = get_internaldata(olle_ptr);
    formatoptions  formatopts = getformatoptions(olle_ptr);
    int            optRowsAtATime = OptRowsAtATime(olle_ptr);
    HRESULT        ret;
    HROW         * row_handle_ptr;
    BOOL           have_hash;
    BOOL           have_array;
    HV           * return_hash;
    AV           * return_array;
    BOOL           new_keys = FALSE;
    SV           * colvalue;

     // Check that we have a active result set.
    if (! mydata->have_resultset) {
        olle_croak (olle_ptr, "Call to nextrow without active result set. Call nextresults first");
    }

    // But the result set may be empty and with out a rowset ptr.
    if (mydata->rowset_ptr != NULL) {

       // If we have row buffer, try to use it.
       if (mydata->rowbuffer != NULL) {
          // We have one, so move to the next row.
          mydata->current_rowno++;

          // But we may now have exhausted the buffer.
          if (mydata->current_rowno > mydata->rows_in_buffer) {
             ret = mydata->rowset_ptr->ReleaseRows(mydata->rows_in_buffer,
                                                   mydata->rowbuffer,
                                                   NULL, NULL, NULL);
             check_for_errors(olle_ptr, "rowset_ptr->ReleaseRows", ret);
             mydata->current_rowno = 0;
             mydata->rows_in_buffer = 0;
             Safefree(mydata->rowbuffer);
             mydata->rowbuffer = NULL;
          }
       }

       // At this point, if we don't have a row buffer, get the first or
       // next one.
       if (mydata->rowbuffer == NULL) {
          New(902, mydata->rowbuffer, optRowsAtATime, HROW);

          // Get rows to the buffer..
          ret = mydata->rowset_ptr->GetNextRows(NULL, 0, optRowsAtATime,
                                                &(mydata->rows_in_buffer),
                                                &(mydata->rowbuffer));
          check_for_errors(olle_ptr, "rowset_ptr->GetNextRows", ret);
          mydata->current_rowno = 1;
       }

       // Now get a pointer, to the current row in the buffer.
       if (mydata->rows_in_buffer > 0) {
          row_handle_ptr = mydata->rowbuffer + (mydata->current_rowno - 1);
       }
       else {
          row_handle_ptr = NULL;
       }
    }
    else {
       row_handle_ptr = NULL;
    }

    // What references did we get?
    have_hash  = (hashref  != NULL && ! SvREADONLY(hashref));
    have_array = (arrayref != NULL && ! SvREADONLY(arrayref));

    if (row_handle_ptr != NULL) {
       // Clear the data buffer to leave room for this row.
       memset(mydata->data_buffer, 0, mydata->size_data_buffer);

       // Get the row data from the rowset.
       ret = mydata->rowset_ptr->GetData(*row_handle_ptr, mydata->row_accessor,
                                         mydata->data_buffer);
       check_for_errors(olle_ptr, "rowset_ptr->GetData", ret);

       // Create the Perl hash and/or array for returning the data.
       if (have_hash) {
          return_hash = newHV();
          sv_setsv(hashref, sv_2mortal(newRV_noinc((SV*) return_hash)));

          // We only determine hash keys once per result set to save some
          // time. Here we allocate an array to save them.
          if (mydata->column_keys == NULL) {
             New(902, mydata->column_keys, mydata->no_of_cols, SV*);
             memset(mydata->column_keys, 0, mydata->no_of_cols * sizeof(SV*));
             new_keys = TRUE;
          }
       }

       if (have_array) {
          return_array = newAV();
          av_extend(return_array, mydata->no_of_cols);
          sv_setsv(arrayref, sv_2mortal(newRV_noinc((SV*) return_array)));
       }

       // Iterate over all columns.
       for (ULONG j = 0; j < mydata->no_of_cols; j++) {
           // Extract the data into colvalue.
           extract_data(olle_ptr, formatopts, FALSE,
                        mydata->column_info[j].pwszName,
                        mydata->column_info[j].wType,
                        mydata->col_bindings[j],
                        mydata->col_bind_status[j],
                        mydata->data_buffer, colvalue);

           // And save the value in the hash.
           if (have_hash) {
              // First get the key as an SV (must be an SV to handle UTF-8
              // correctly. Allocate one on first round and save.
              if (new_keys) {
                 SV * colkey;

                 if (wcslen(mydata->column_info[j].pwszName) > 0) {
                    // There is a column name, lets use it.
                    colkey = BSTR_to_SV(mydata->column_info[j].pwszName);
                 }
                 else {
                    // Anonymous column, construct a default name.
                    char  tmp[20];
                    sprintf(tmp, "Col %d", j + 1);
                    colkey = newSVpv(tmp, strlen(tmp));
                 }

                 // Check for duplicates and iterate till we have one, but
                 // we don't try forever.
                 char c = '@';
                 while (hv_exists_ent(return_hash, colkey, 0) && c++ <= 'Z') {
                    if (PL_dowarn) {
                       warn("Column name '%s' appears twice or more in the result set",
                            SvPV_nolen(colkey));
                    }
                    SvREFCNT_dec(colkey);
                    char  tmp[20];
                    sprintf(tmp, "Col %d%c", j + 1, c);
                    colkey = newSVpv(tmp, strlen(tmp));
                 }

                 // Save the key value.
                 mydata->column_keys[j] = colkey;
              }

              // And now store the column value.
              hv_store_ent(return_hash, mydata->column_keys[j], colvalue, 0);
           }

           // And save to the array. Note that if we save in both hash and
           // array, we need to bump the reference count.
           if (have_array) {
              if (have_hash) {
                 SvREFCNT_inc(colvalue);
              }
              av_store(return_array, j, colvalue);
           }
       }
    }
    else {
       // Last row in result set. Set return references to undef, and free
       // up memory for the result set.
       free_resultset_data(mydata);
       if (have_hash) {
          sv_setsv(hashref, &PL_sv_undef);
       }
       if (have_array) {
          sv_setsv(arrayref, &PL_sv_undef);
       }
   }

   // Set the rerurn value.
   return (row_handle_ptr != NULL ? 1 : 0);
}

//-----------------------------------------------------------------------
// $X->getoutputparams.
//-----------------------------------------------------------------------
void getoutputparams (SV * olle_ptr,
                      SV * hashref,
                      SV * arrayref)
{
    internaldata  * mydata = get_internaldata(olle_ptr);
    formatoptions   formatopts = getformatoptions(olle_ptr);
    paramdata     * current_param;
    ULONG           parno = 0;
    ULONG           outparno = 0;
    BOOL            have_hash;
    BOOL            have_array;
    HV            * return_hash;
    AV            * return_array;
    SV            * parvalue;

    // Check that we have a active result set.
    if (mydata->no_of_out_params == 0) {
        olle_croak (olle_ptr, "Call to getoutputparams for a batch that did not have output parameters");
    }

    if (! mydata->params_available) {
       olle_croak(olle_ptr, "Output parameters are not available at this point. First get all results sets with nextresultset");
    }

    // What references did we get?
    have_hash  = (hashref  != NULL && ! SvREADONLY(hashref));
    have_array = (arrayref != NULL && ! SvREADONLY(arrayref));

    // Create the Perl hash and/or array for returning the data.
    if (have_hash) {
       return_hash = newHV();
       sv_setsv(hashref, sv_2mortal(newRV_noinc((SV*) return_hash)));
    }
    if (have_array) {
       return_array = newAV();
       av_extend(return_array, mydata->no_of_cols);
       sv_setsv(arrayref, sv_2mortal(newRV_noinc((SV*) return_array)));
    }

    // Iterate over all parameters.
    current_param = mydata->paramfirst;
    while (current_param != NULL) {
       parno++;

       // But only output parameters are interesting.
       if (current_param->isoutput) {
          outparno++;

          // Extract the data into paravalue.
          extract_data(olle_ptr, formatopts, TRUE,
                       mydata->param_info[parno - 1].pwszName,
                       current_param->datatype,
                       mydata->param_bindings[parno - 1],
                       mydata->param_bind_status[parno - 1],
                       mydata->param_buffer, parvalue);
          // And save the value in the hash and/or array.
          if (have_hash) {
             // Need to construct a key first. It must be an SV to get
             // UTF-8 right.
             SV * hashkey;
             if (current_param->param_info.pwszName != NULL &&
                 wcslen(current_param->param_info.pwszName) > 0) {
                hashkey = BSTR_to_SV(current_param->param_info.pwszName);
             }
             else {
                char tmp[20];
                sprintf(tmp, "Par %d", outparno);
                hashkey = newSVpv(tmp, strlen(tmp));
             }
             hv_store_ent(return_hash, hashkey, parvalue, 0);
             SvREFCNT_dec(hashkey);
          }

          if (have_array) {
             if (have_hash) {
                SvREFCNT_inc(parvalue);
             }
             av_store(return_array, outparno - 1, parvalue);
          }
       }

       // Move to next.
       current_param = current_param->next;
    }

    // The batch is now completely exhausted, so we can free all resources
    // bound to it.
    free_batch_data(mydata);
}


//======================================================================
// The XS part of it all.
// Most routines are just declarations.
//======================================================================
MODULE = Win32::SqlServer           PACKAGE = Win32::SqlServer

PROTOTYPES: ENABLE

BOOT:
initialize();

void
olledb_message (olle_ptr, msgno, state, severity, msg)
   SV   * olle_ptr
   int    msgno
   int    state
   int    severity
   char * msg

int
setupinternaldata()

void
setloginproperty(olle_ptr, prop_name, prop_value)
   SV   * olle_ptr;
   char * prop_name;
   SV   * prop_value;


int
connect(olle_ptr)
   SV * olle_ptr
  CODE:
{
    internaldata  * mydata = get_internaldata(olle_ptr);

    // Check that we are not already connected.
    if (mydata->datasrc_ptr != NULL) {
       olle_croak(olle_ptr, "Attempt to connect despite already being connected");
    }

    RETVAL = do_connect(olle_ptr, FALSE);
}
OUTPUT:
   RETVAL

void
disconnect(olle_ptr)
   SV * olle_ptr

int
isconnected(olle_ptr)
   SV * olle_ptr
  CODE:
{
   internaldata  * mydata = get_internaldata(olle_ptr);
   RETVAL = mydata->datasrc_ptr != NULL;
}
OUTPUT:
   RETVAL

void
xs_DESTROY(olle_ptr)
        SV *    olle_ptr
  CODE:
{
// This routine is called from DESTROY in the Perl code. We cannot have
// DESTROY here directly, because the Perl code has to take some extra
// precautions.

    internaldata * mydata = get_internaldata(olle_ptr);

    if (mydata != NULL) {
       disconnect(olle_ptr);

       // Free up area allocated to all properties.
       for (int i = 0; gbl_init_props[i].propset_enum != not_in_use; i++) {
          VariantClear(&mydata->init_properties[i].vValue);
       }

       // And dispense of mydata itself. The Perl DESTROY will set mydata
       // to 0, to avoid a second cleanup when Perl calls DESTROY a second
       // time. (Which it does for some reason.)
       Safefree(mydata);
   }
}

void
validatecallback(olle_ptr, callbackname)
          SV * olle_ptr
          SV * callbackname
CODE:
{
    // This is a help routine to validate that a name for a message handler
    // refers to an existing sub. It's called from STORE (which is in Perl
    // code).
    char *name = SvPV_nolen(callbackname);
    CV * callback = get_cv(name, FALSE);
    if (! callback) {
        olle_croak(olle_ptr, "Can't find specified message handler '%s'", name);
    }
    // OK, we found an message handler, but was it pure luck?
    else if (PL_dowarn && ! strstr(name, "::")) {
       warn("Message handler '%s' given as a unqualified name. This could fail next time you try", name);
    }
}

void
initbatch(olle_ptr, sv_cmdtext)
    SV  *olle_ptr
    SV  *sv_cmdtext

int
enterparameter(olle_ptr, nameoftype, sv_maxlen, paramname, isinput, isoutput, sv_value = NULL, precision = 18, scale = 0, typeinfo = NULL)
   SV            * olle_ptr
   SV            * nameoftype
   SV            * sv_maxlen
   SV            * paramname
   int             isinput
   int             isoutput
   SV            * sv_value
   unsigned char   precision
   unsigned char   scale
   SV            * typeinfo

int
executebatch(olle_ptr, rows_affected = NULL)
  SV * olle_ptr;
  SV * rows_affected;

int
nextresultset(olle_ptr, rows_affected = NULL)
  SV * olle_ptr;
  SV * rows_affected;

int
nextrow (olle_ptr, hashref, arrayref)
    SV * olle_ptr
    SV * hashref
    SV * arrayref
OUTPUT:
   RETVAL
   hashref
   arrayref

void
getoutputparams (olle_ptr, hashref, arrayref)
    SV * olle_ptr
    SV * hashref
    SV * arrayref
OUTPUT:
   hashref
   arrayref


void
cancelbatch (olle_ptr)
    SV * olle_ptr
CODE:
{
    internaldata * mydata = get_internaldata(olle_ptr);
    free_batch_data(mydata);
}

void
cancelresultset (olle_ptr)
    SV * olle_ptr
CODE:
{
    internaldata * mydata = get_internaldata(olle_ptr);
    free_resultset_data(mydata);
}

int
getcmdstate (olle_ptr)
    SV * olle_ptr
CODE:
{
    internaldata * mydata = get_internaldata(olle_ptr);

    if (mydata->pending_cmd == NULL) {
       RETVAL = cmdstate_init;
    }
    else if (mydata->cmdtext_ptr == NULL) {
       RETVAL = cmdstate_enterexec;
    }
    else if (mydata->params_available) {
       RETVAL = cmdstate_getparams;
    }
    else if (mydata->have_resultset) {
       RETVAL = cmdstate_nextrow;
    }
    else {
       RETVAL = cmdstate_nextres;
    }
}
OUTPUT:
   RETVAL

SV *
getcmdtext (olle_ptr)
    SV * olle_ptr
CODE:
{
    internaldata * mydata = get_internaldata(olle_ptr);
    if (mydata->pending_cmd != NULL) {
       RETVAL = BSTR_to_SV(mydata->pending_cmd);
    }
    else {
       RETVAL = &PL_sv_undef;
    }
}
OUTPUT:
   RETVAL

int
get_provider_enum(olle_ptr)
    SV * olle_ptr
CODE:
{
    // Implements FETCH for Olle->{Provider}.
    internaldata * mydata = get_internaldata(olle_ptr);
    RETVAL = mydata->provider;
}
OUTPUT:
   RETVAL

int
set_provider_enum(olle_ptr, provider)
    SV * olle_ptr
    int  provider;
CODE:
{
    // Implements STORE for Olle->{Provider}. We return -1 if connected.
    // The Perl module will do the croaking for better location of error
    // message.
    internaldata * mydata = get_internaldata(olle_ptr);
    if (mydata->datasrc_ptr != NULL) {
       RETVAL = -1;
    }
    else {
       mydata->provider = (provider_enum) provider;
       if (mydata->provider == provider_default) {
          // If unknown, take SQLNCLI if it's available.
          mydata->provider = (IsEqualCLSID(clsid_sqlncli, CLSID_NULL) ?
                               provider_sqloledb : provider_sqlncli);
       }
       RETVAL = mydata->provider;
    }
}
OUTPUT:
   RETVAL


void
parsename(olle_ptr, sv_namestr, retain_quotes, sv_server, sv_db, sv_schema, sv_object)
   SV * olle_ptr
   SV * sv_namestr
   int retain_quotes
   SV * sv_server
   SV * sv_db
   SV * sv_schema
   SV * sv_object

void
replaceparamholders (olle_ptr, cmdstring)
   SV * olle_ptr
   SV * cmdstring
