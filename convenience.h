/*---------------------------------------------------------------------
 $Header: /Perl/OlleDB/convenience.h 2     08-02-24 21:59 Sommar $

  This file holds general-purpose routines, mainly for converting
  between SV and BSTR and the like. All these are low-level, and do
  not have access to error handling. Such code should be in utils.cpp.

  Copyright (c) 2004-2008   Erland Sommarskog

  $History: convenience.h $
 * 
 * *****************  Version 2  *****************
 * User: Sommar       Date: 08-02-24   Time: 21:59
 * Updated in $/Perl/OlleDB
 * Added quotename().
 *
 * *****************  Version 1  *****************
 * User: Sommar       Date: 07-12-24   Time: 21:39
 * Created in $/Perl/OlleDB
  ---------------------------------------------------------------------*/


extern BSTR char_to_BSTR(char     * str,
                         STRLEN     inlen,
                         BOOL       isutf8,
                         DBLENGTH * bytelen = NULL,
                         BOOL       add_BOM = FALSE);

extern BSTR SV_to_BSTR (SV       * sv,
                        DBLENGTH * bytelen = NULL,
                        BOOL       add_BOM = FALSE);

extern char * BSTR_to_char (BSTR bstr);

extern SV * BSTR_to_SV (BSTR  bstr,
                       int   bstrlen = -1);

extern LONG pow10(unsigned int n);

extern void quotename(BSTR &str);