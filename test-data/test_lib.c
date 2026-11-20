/*
 * Tests for MAPISysTimetoDTR(), to_utf8(), MAPIPrint() on a value with no
 * data, the DecompressRTF() output size limit, and short reads in
 * TNEFParseMemory().
 */
#define _DEFAULT_SOURCE
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include <ytnef.h>

static int failed = 0;

static void check(int ok, const char *what) {
  printf("%s: %s\n", ok ? "PASS" : "FAIL", what);
  if (!ok) failed++;
}

static void date(int y, int m, int d, int dow) {
  struct tm t = {0};
  time_t u;
  unsigned long long ft;
  BYTE buf[8];
  dtr out;
  char name[64];

  t.tm_year = y - 1900;
  t.tm_mon = m - 1;
  t.tm_mday = d;
  t.tm_hour = 13;
  t.tm_min = 45;
  u = timegm(&t);
  ft = (unsigned long long)((long long)u + 11644473600LL) * 10000000ULL;
  memcpy(buf, &ft, sizeof(buf));
  MAPISysTimetoDTR(buf, &out);
  snprintf(name, sizeof(name), "date %04d-%02d-%02d", y, m, d);
  check(out.wYear == y && out.wMonth == m && out.wDay == d &&
        out.wDayOfWeek == dow && out.wHour == 13 && out.wMinute == 45, name);
}

static void utf8(const char *utf16le, size_t len, const char *expect,
                 const char *what) {
  char *s = to_utf8(len, (char *)utf16le);
  check(s != NULL && strcmp(s, expect) == 0, what);
  free(s);
}

int main(void) {
  BYTE short_time[4] = {0};
  dtr out;

  date(2003, 2, 1, 6);
  date(2003, 3, 1, 6);
  date(2002, 10, 1, 2);
  date(2000, 2, 29, 2);
  date(2000, 3, 1, 3);
  date(1900, 3, 1, 4);
  date(2004, 1, 1, 4);
  date(2024, 12, 31, 2);
  check(MAPISysTimetoDTRSize(short_time, sizeof(short_time), &out) == -1,
        "date shorter than 8 bytes");

  utf8("A\0", 2, "A", "utf8 U+0041");
  utf8("\xff\x07", 2, "\xdf\xbf", "utf8 U+07FF");
  utf8("\x00\x08", 2, "\xe0\xa0\x80", "utf8 U+0800");
  utf8("\x3d\xd8\x00\xde", 4, "\xf0\x9f\x98\x80", "utf8 surrogate pair");
  utf8("\x00\xd8" "A\0", 4, "\xef\xbf\xbd" "A", "utf8 lone surrogate");
  utf8("", 0, "", "utf8 empty");

  {
    variableLength value = { NULL, 0 };
    MAPIProperty prop;
    MAPIProps props;

    memset(&prop, 0, sizeof(prop));
    prop.id = PROP_TAG(PT_STRING8, 0x0050);
    prop.count = 1;
    prop.data = &value;
    props.count = 1;
    props.properties = &prop;
    fflush(stdout);
    MAPIPrint(&props);
    check(1, "MAPIPrint of a value with no data");
  }

  {
    /* compressed stream that claims 2 GiB of output from 20 bytes */
    BYTE raw[20] = {
      16, 0, 0, 0,  0, 0, 0, 0x7f,  0x4c, 0x5a, 0x46, 0x75,  0, 0, 0, 0,
      0, 'a', 'b', 'c'
    };
    variableLength p;
    int size = 0;
    BYTE *r;

    p.size = sizeof(raw);
    p.data = malloc(p.size);
    memcpy(p.data, raw, p.size);
    r = DecompressRTF(&p, &size);
    check(r == NULL, "DecompressRTF refuses an impossible size");
    free(r);
    free(p.data);
  }

  {
    /* same bytes as regress/truncated-field.tnef: a field that claims
     * 4 bytes of data when only 3 remain */
    BYTE raw[18] = {
      0x78, 0x9f, 0x3e, 0x22,  0x01, 0x00,
      0x01,  0x06, 0x90, 0x08, 0x00,  0x04, 0x00, 0x00, 0x00,
      0x00, 0x00, 0x00
    };
    TNEFStruct t;

    TNEFInitialize(&t);
    check(TNEFParseMemory(raw, sizeof(raw), &t) < 0,
          "TNEFParseMemory rejects a truncated field");
    TNEFFree(&t);
  }

  {
    /* issue #109: a named property whose UTF-16 name is 114 bytes long */
    BYTE raw[6 + 9 + 152 + 2];
    BYTE *m = raw + 15;
    DWORD v;
    WORD sum = 0;
    int i;
    TNEFStruct t;

    memset(raw, 0, sizeof(raw));
    v = 0x223e9f78; memcpy(raw, &v, 4);          /* signature */
    raw[6] = 0x01;                                /* message level */
    v = attMAPIProps; memcpy(raw + 7, &v, 4);
    v = 152; memcpy(raw + 11, &v, 4);             /* field size */
    v = 1; memcpy(m, &v, 4);                      /* property count */
    v = 0x80000003; memcpy(m + 4, &v, 4);         /* PT_LONG, named */
    v = 1; memcpy(m + 24, &v, 4);                 /* one name */
    v = 114; memcpy(m + 28, &v, 4);               /* name size */
    for (i = 0; i < 114; i += 2) m[32 + i] = 'a';
    v = 42; memcpy(m + 148, &v, 4);               /* value, after padding */
    for (i = 0; i < 152; i++) sum += m[i];
    memcpy(m + 152, &sum, 2);

    TNEFInitialize(&t);
    check(TNEFParseMemory(raw, sizeof(raw), &t) == 0,
          "TNEFParseMemory accepts a 57-character property name");
    check(t.MapiProperties.count == 1 &&
          t.MapiProperties.properties[0].propnames[0].size == 114,
          "long property name is kept");
    TNEFFree(&t);
  }

  return failed ? 1 : 0;
}
