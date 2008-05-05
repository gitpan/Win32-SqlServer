/*---------------------------------------------------------------------
 $Header: /Perl/OlleDB/init.cpp 3     08-04-30 22:46 Sommar $

  This file holds code that is run when the module initialiases, and
  when a new OlleDB object is created. This file also declares global
  variables that exist through the lifetime of the module. They are
  constants that are set up once and then never changed.


  Copyright (c) 2004-2008   Erland Sommarskog

  $History: init.cpp $
 * 
 * *****************  Version 3  *****************
 * User: Sommar       Date: 08-04-30   Time: 22:46
 * Updated in $/Perl/OlleDB
 * Use get_sv and not perl_get_sv (deprecated). Pass GV_ADDMULTI to get_sv
 * to avoid "Used only once" warning. Don't define macro XS_VERSION in the
 * file, as it comes with the Makefile.
 *
 * *****************  Version 2  *****************
 * User: Sommar       Date: 08-01-06   Time: 23:33
 * Updated in $/Perl/OlleDB
 * Replaced all unsafe CRT functions with their safe replacements in VC8.
 * olledb_message now takes a va_list as argument, so we pass it
 * parameterised strings and don't have to litter the rest of the code
 * with that.
 *
 * *****************  Version 1  *****************
 * User: Sommar       Date: 07-12-24   Time: 21:40
 * Created in $/Perl/OlleDB
  ---------------------------------------------------------------------*/

#define _WIN32_DCOM   // Needed for CoInitializeEx

#include "CommonInclude.h"

#include <cguid.h>
#include <msdaguid.h>


#include "convenience.h"
#include "datatypemap.h"
#include "init.h"



#undef FILEDEBUG
#ifdef FILEDEBUG
FILE *dbgfile = NULL;
#endif


// Global variables for class ids for the possible providers.
CLSID  clsid_sqloledb  = CLSID_NULL;
CLSID  clsid_sqlncli   = CLSID_NULL;
CLSID  clsid_sqlncli10 = CLSID_NULL;

// This global array holds definition of all initialisation properties
// for OLE DB.
init_property gbl_init_props[MAX_INIT_PROPERTIES];

// This global holds how many of the SSINIT properties that applies to
// SQLOLEDB - there are some that only applies to SQL Native Client.
int no_of_sqloledb_ssprops;

// This array holds where each property set starts in gbl_init_props;
propset_info_struct init_propset_info[NO_OF_INIT_PROPSETS];


// Global pointer to OLE DB Services. Set once when we intialize, and
// never released.
IDataInitialize * data_init_ptr    = NULL;

// Global pointer the OLE DB conversion library.
IDataConvert    * data_convert_ptr = NULL;

// Global pointer to the IMalloc interface. Most of the time when we allocate
// memory, we rely on the Perl methods. However, there are situations when
// we must free memory allocated by SQLOLEDB. Same here, we create once, as
// the COM implementation is touted as thread-safe.
IMalloc*   OLE_malloc_ptr = NULL;



// A helper routine to get default for APPNAME.
static BSTR get_scriptname () {
   // Get the name of the script, taken from Perl var $0. This is used as
   // the default application name in SQL Server.

   SV* sv;

   if (sv = get_sv("0", FALSE))
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

   strcpy_s(gbl_init_props[ix].name, INIT_PROPNAME_LEN, name);
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
   char      * obj;

   // In the critical section we create our starting point, the pointer to
   // OLE DB services. We also create a pointer to a conversion object.
   // Thess pointer will never be released.
   EnterCriticalSection(&CS);

   // Get classIDs for the possible providers.
   if (IsEqualCLSID(clsid_sqloledb, CLSID_NULL) &&
       IsEqualCLSID(clsid_sqlncli, CLSID_NULL)  &&
       IsEqualCLSID(clsid_sqlncli10, CLSID_NULL)) {

      ret = CLSIDFromProgID(L"SQLOLEDB", &clsid_sqloledb);
      if (FAILED(ret)) {
         clsid_sqloledb = CLSID_NULL;
      }

      ret = CLSIDFromProgID(L"SQLNCLI", &clsid_sqlncli);
      if (FAILED(ret)) {
         clsid_sqlncli = CLSID_NULL;
      }

      ret = CLSIDFromProgID(L"SQLNCLI10", &clsid_sqlncli10);
      if (FAILED(ret)) {
         clsid_sqlncli10 = CLSID_NULL;
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
         obj = "IDataInitialize";
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
         obj = "IDataConvert";
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
   if (sv = get_sv("Win32::SqlServer::Version", GV_ADD | GV_ADDMULTI))
   {
        char buff[256];
        sprintf_s(buff, 256,
                  "This is Win32::SqlServer, version %s\n\nCopyright (c) 2005-2008 Erland Sommarskog\n",
                  XS_VERSION);
        sv_setnv(sv, atof(XS_VERSION));
        sv_setpv(sv, buff);
        SvNOK_on(sv);
   }
}


// This routine returns the default provider, which is highest version of
// SQL Native Client/SQLOLEDB that is installed.
provider_enum default_provider(void) {
  if (! IsEqualCLSID(clsid_sqlncli10, CLSID_NULL))
      return provider_sqlncli10;
  else if (! IsEqualCLSID(clsid_sqlncli, CLSID_NULL))
      return provider_sqlncli;
  else
      return provider_sqloledb;
}

