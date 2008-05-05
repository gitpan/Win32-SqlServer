/*---------------------------------------------------------------------
 $Header: /Perl/OlleDB/utils.cpp 3     08-05-01 23:54 Sommar $

  This file includes various utility routines. In difference to
  the convenience routines, these may call the error handler and
  that. Several of these are called from Perl code as well.

  Copyright (c) 2004-2008   Erland Sommarskog

  $History: utils.cpp $
 * 
 * *****************  Version 3  *****************
 * User: Sommar       Date: 08-05-01   Time: 23:54
 * Updated in $/Perl/OlleDB
 * Rewrote codepage_convert so that the new text is written to a new
 * buffer, and not directly into the old, as that could cause problems
 * with shared hash keys.
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

#include "CommonInclude.h"
#include "convenience.h"
#include "errcheck.h"
#include "utils.h"


//---------------------------------------------------------------------
// This is a helper routine called from get_object_id in the Perl code.
// It cracks an object specification into its parts, retaining any quotes
// around the identifiers and returns the result in sv_server, sv_db,
// sv_schema and sv_object.
//------------------------------------------------------------------
void parsename(SV   * olle_ptr,
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
// This is a helper routine that scans a SQL command string for ? and
// replaces them with @P1, @P2 etc. It's called from setup_sqlcommand
// in the Perl module, and not called by the XS code. It's written in C++,
// simply because it appeared simpler than to do it in Perl.
//----------------------------------------------------------------------
void replaceparamholders (SV * olle_ptr,
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
            sprintf_s(paramstr, 12, "@P%d", parno++);
            strcpy_s(&(output[outix]), 3*inputlen + 3 - outix, paramstr);
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


//-----------------------------------------------------------------------------
// change_codepage, used to implement sql_set_conversion.
//-----------------------------------------------------------------------------
void codepage_convert(SV     * olle_ptr,
                      SV     * sv,
                      UINT     from_cp,
                      UINT     to_cp)

{  int      widelen;
   int      ret;
   DWORD    err;
   BSTR     bstr;
   STRLEN   inputlen;
   char   * input;
   STRLEN   outlen;
   char   * output;

   // Get out if this is not a string.
   if (! SvPOK(sv)) return;

   input = SvPV(sv, inputlen);

   // And if the string has no length.
   if (inputlen == 0) return;

   // If the input string is UTF_8, we should ignore from_cp.
   if (SvUTF8(sv)) {
      from_cp = CP_UTF8;
   }

   // First find out how long the Unicode string will be, by calling
   // MultiByteToWideChar without a buffer. Not that we always set flags to
   // 0 here, since it works with all code pages.
   widelen = MultiByteToWideChar(from_cp, 0, input, inputlen, NULL, 0);

   if (widelen > 0) {
      // Allocate Unicode string and convert to Unicode.
      bstr = SysAllocStringLen(NULL, widelen);
      ret = MultiByteToWideChar(from_cp, 0, input, inputlen, bstr, widelen);
   }
   else {
      ret = 0;
   }

   // Check for errors.
   if (ret == 0) {
      err = GetLastError();
      if (err == ERROR_INVALID_PARAMETER) {
         olle_croak(olle_ptr,
                    "Conversion from codepage %d to Unicode failed. Maybe you are using an non-existing code-page?",
                    from_cp);
      }
      else {
         olle_croak(olle_ptr,
                    "Conversion from codepage %d to Unicode failed with error %d",
                    from_cp, err);
      }
   }

   // Now determine the length for the string in the receiving code page.
   outlen = WideCharToMultiByte(to_cp, 0, bstr, widelen, NULL, 0, NULL, NULL);

   if (outlen > 0) {
      New(902, output, outlen, char);
      ret = WideCharToMultiByte(to_cp, 0, bstr, widelen, output, outlen, NULL, NULL);
   }
   else {
      ret = 0;
   }

   if (ret == 0) {
      err = GetLastError();
      if (err == ERROR_INVALID_PARAMETER) {
         olle_croak(olle_ptr,
                    "Conversion to codepage %d from Unicode failed. Maybe you are using an non-existing code-page?",
                    to_cp);
      }
      else {
         olle_croak(olle_ptr,
                    "Conversion to codepage %d from Unicode failed with error %d",
                    to_cp, err);
      }
   }

   // Copy into the SV.
   sv_setpvn(sv, output, outlen);

   // Get rid of the bstr and the output buffer.
   SysFreeString(bstr);
   Safefree(output);

   // Set or unset the UTF8 flag depending on target charset.
   if (to_cp == CP_UTF8) {
      SvUTF8_on(sv);
   }
   else {
      SvUTF8_off(sv);
   }
}

