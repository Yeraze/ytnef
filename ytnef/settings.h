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

#ifndef YTNEF_SRC_YTNEF_SETTINGS_H_
#define YTNEF_SRC_YTNEF_SETTINGS_H_

#define MAX_FILENAME_SIZE 1024

void SanitizeFilename(char *filename);
void CreateUniqueFilename(char* output, unsigned int max_size,
                          const char* name, const char* ext,
                          const char* path);

#endif  // YTNEF_SRC_YTNEF_SETTINGS_H_
