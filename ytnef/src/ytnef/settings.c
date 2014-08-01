/*
*    Yerase's TNEF Stream Reader
*    Copyright (C) 2003  Randall E. Hand
*
*    This program is free software; you can redistribute it and/or modify
*    it under the terms of the GNU General Public License as published by
*    the Free Software Foundation; either version 2 of the License, or
*    (at your option) any later version.
*
*    This program is distributed in the hope that it will be useful,
*    but WITHOUT ANY WARRANTY; without even the implied warranty of
*    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
*    GNU General Public License for more details.
*
*    You should have received a copy of the GNU General Public License
*    along with this program; if not, write to the Free Software
*    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
*
*    You can contact me at randall.hand@gmail.com for questions or assistance
*/

#include "settings.h"


// Replace every character in a filename (in place)
// that is not a valid AlphaNumeric (a-z, A-Z, 0-9) or a period
// with an underscore.
void SanitizeFilename(char *filename) {
  int i;
  for (i = 0; i < strlen(filename); ++i) {
    if (! (isalnum(filename[i]) || (filename[i] == '.'))) {
      filename[i] = '_';
    }
  }
}

