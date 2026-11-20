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
// with an underscore.  Directory separators are replaced too.
void SanitizeFilename(char *filename) {
  int i;
  for (i = 0; i < strlen(filename); ++i) {
    if (! (isalnum((unsigned char)filename[i]) || (filename[i] == '.'))) {
      filename[i] = '_';
    }
  }
}

// Make a TNEF-supplied attachment name safe to use as a single path
// component: spaces and directory separators become underscores, and
// names that would still refer to a directory ("", ".", "..") are
// rejected.  Returns 0 if the name is unusable.
int SanitizeAttachmentName(char *filename) {
  int i;
  for (i = 0; filename[i] != 0; ++i) {
    if (filename[i] == ' ' || filename[i] == '/' || filename[i] == '\\') {
      filename[i] = '_';
    }
  }
  if (filename[0] == 0 || strcmp(filename, ".") == 0 ||
      strcmp(filename, "..") == 0) {
    return 0;
  }
  return 1;
}

void CreateUniqueFilename(char* output, unsigned int max_size,
                          const char* name, const char* ext,
                          const char* path) {
  int counter = 0;
  FILE* check = NULL;
  while (1) {
    if (filepath != NULL) {
      if (counter > 0){
        snprintf(output, max_size, "%s/%s%i.%s", path, name, counter, ext);
      } else {
        snprintf(output, max_size, "%s/%s.%s", path, name, ext);
      }
    } else {
      if (counter > 0){
        snprintf(output, max_size, "%s%i.%s", name, counter, ext);
      } else {
        snprintf(output, max_size, "%s.%s", name, ext);
      }
    }
    //Let's check if there is a file with the same name.
    if ((check = fopen(output, "r")) == NULL) {
      // There was an error, so this would be unique
      break;
    } else {
      // We successfully opened the file, so close 
      // and try again
      fclose(check);
      counter++;
    }
  }
  return;
}
