/*---------------------------------------------------------------------
 $Header: /Perl/OlleDB/utils.h 1     07-12-24 21:39 Sommar $

  This file includes various utility routines. In difference to
  the convenience routines, these may call the error handler and
  that. Several of these are called from Perl code as well.

  Copyright (c) 2004-2008   Erland Sommarskog

  $History: utils.h $
 * 
 * *****************  Version 1  *****************
 * User: Sommar       Date: 07-12-24   Time: 21:39
 * Created in $/Perl/OlleDB
  ---------------------------------------------------------------------*/


extern void parsename(SV   * olle_ptr,
                      SV   * sv_namestr,
                      int    retain_quotes,
                      SV   * sv_server,
                      SV   * sv_db,
                      SV   * sv_schema,
                      SV   * sv_object);


extern void replaceparamholders (SV * olle_ptr,
                                SV * cmdstring);

extern void codepage_convert(SV     * olle_ptr,
                             SV     * sv,
                             UINT     from_cp,
                             UINT     to_cp);

