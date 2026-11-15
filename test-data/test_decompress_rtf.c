/*
 * Tests for DecompressRTF(), IsCompressedRTF() and the data of a MAPI property
 * when parsing stops part way through it.
 *
 * Each case runs in a child process so that one failure, or a sanitizer stop,
 * does not hide the others. Streams are in exactly sized heap buffers so a
 * read past the end is visible under AddressSanitizer.
 */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/wait.h>
#include <unistd.h>
#include <ytnef.h>

#define MAGIC_UNCOMPRESSED 0x414c454d
#define MAGIC_COMPRESSED   0x75465a4c

int IsCompressedRTF(variableLength *p);

static void put32(BYTE *b, unsigned int v) {
  b[0] = v & 0xff;
  b[1] = (v >> 8) & 0xff;
  b[2] = (v >> 16) & 0xff;
  b[3] = (v >> 24) & 0xff;
}

/* A stream: 16-byte header, then the payload. */
static variableLength stream(unsigned int magic, unsigned int rawsize,
                             const char *payload, int len) {
  variableLength p;

  p.size = 16 + len;
  p.data = calloc(p.size, 1);
  put32(p.data, p.size - 4);
  put32(p.data + 4, rawsize);
  put32(p.data + 8, magic);
  memcpy(p.data + 16, payload, len);
  return p;
}

static int refused(variableLength *p) {
  int out = -1;
  BYTE *r = DecompressRTF(p, &out);
  int ok = (r == NULL);

  free(r);
  free(p->data);
  return ok ? 0 : 1;
}

static int gives(variableLength *p, const char *want) {
  int out = -1;
  BYTE *r = DecompressRTF(p, &out);
  int ok = (r != NULL) && (out == (int)strlen(want))
           && (memcmp(r, want, out) == 0) && (r[out] == '\0');

  free(r);
  free(p->data);
  return ok ? 0 : 1;
}

static int no_data(void) {
  variableLength p;
  int out = 0;

  p.data = NULL;
  p.size = 0;
  if (DecompressRTF(&p, &out) != NULL) return 1;
  p.size = 20;
  if (DecompressRTF(&p, &out) != NULL) return 1;
  p.size = 0x2b0000;
  if (DecompressRTF(&p, &out) != NULL) return 1;
  return 0;
}

static int too_short(void) {
  BYTE shortbuf[4] = { 0, 0, 0, 0 };
  variableLength p;
  int out = 0;

  p.data = shortbuf;
  p.size = sizeof(shortbuf);
  return DecompressRTF(&p, &out) != NULL;
}

static int raw_good(void) {
  variableLength p = stream(MAGIC_UNCOMPRESSED, 5, "hello", 5);
  return gives(&p, "hello");
}

static int raw_claims_too_much(void) {
  variableLength p = stream(MAGIC_UNCOMPRESSED, 0x4000, "hello", 5);
  return refused(&p);
}

static int compressed_good(void) {
  /* flag byte: all literals */
  variableLength p = stream(MAGIC_COMPRESSED, 3, "\0abcd", 5);
  return gives(&p, "abc");
}

static int compressed_one_byte_for_reference(void) {
  /* flag byte: three literals, then a reference with one of its bytes left */
  variableLength p = stream(MAGIC_COMPRESSED, 100, "\x08" "abcd", 5);
  return refused(&p);
}

static int compressed_no_bytes_for_reference(void) {
  /* eight literals, then a flag byte is the last byte: a reference follows it */
  variableLength p = stream(MAGIC_COMPRESSED, 100, "\0abcdefgh\x01", 10);
  return refused(&p);
}

static int compressed_no_bytes_for_literal(void) {
  /* eight literals, then a flag byte is the last byte: a literal follows it */
  variableLength p = stream(MAGIC_COMPRESSED, 100, "\0abcdefgh\0", 10);
  return refused(&p);
}

static int compressed_huge_size(void) {
  variableLength p = stream(MAGIC_COMPRESSED, 0xffffffffU, "\0abcd", 5);
  return refused(&p);
}

static int unknown_magic(void) {
  variableLength p = stream(0x12345678, 5, "hello", 5);
  return refused(&p);
}

static int is_compressed(void) {
  variableLength p;
  int bad = 0;

  p.data = NULL;
  p.size = 64;
  bad += (IsCompressedRTF(&p) != 0);
  p.size = 0;
  bad += (IsCompressedRTF(&p) != 0);
  p = stream(MAGIC_COMPRESSED, 3, "\0abcd", 5);
  bad += (IsCompressedRTF(&p) != 1);
  free(p.data);
  return bad;
}

/* A property whose stated size is larger than the stream must not be left
 * with that size and no data once parsing has failed. */
static int parse_stops_in_property(void) {
  static const unsigned int tag = 0x00069003; /* attMAPIProps */
  BYTE props[16];
  BYTE file[6 + 9 + sizeof(props) + 2];
  BYTE *d = file;
  TNEFStruct t;
  variableLength *v;
  unsigned int sum = 0;
  size_t i;
  int bad = 0;

  put32(props, 1);                                       /* property count */
  put32(props + 4, PROP_TAG(PT_BINARY, PR_RTF_COMPRESSED));
  put32(props + 8, 1);                                   /* value count */
  put32(props + 12, 0x2b0000);                           /* value size */
  put32(d, 0x223e9f78); d += 4;                          /* signature */
  *d++ = 1; *d++ = 0;                                    /* key */
  *d++ = 1;                                              /* level */
  put32(d, tag); d += 4;
  put32(d, sizeof(props)); d += 4;
  memcpy(d, props, sizeof(props)); d += sizeof(props);
  for (i = 0; i < sizeof(props); i++) sum += props[i];
  *d++ = sum & 0xff;
  *d++ = (sum >> 8) & 0xff;

  TNEFInitialize(&t);
  if (TNEFParseMemory(file, sizeof(file), &t) >= 0) bad++;
  v = MAPIFindProperty(&t.MapiProperties, PROP_TAG(PT_BINARY, PR_RTF_COMPRESSED));
  if (v != MAPI_UNDEFINED && v->data == NULL && v->size != 0) bad++;
  TNEFFree(&t);
  return bad;
}

static int string_without_nul(void) {
  static const unsigned int tag = 0x00069003; /* attMAPIProps */
  BYTE props[16 + 8];
  BYTE file[6 + 9 + sizeof(props) + 2];
  BYTE *d = file;
  TNEFStruct t;
  variableLength *v;
  unsigned int sum = 0;
  size_t i;
  int bad = 0;

  memset(props, 0, sizeof(props));
  put32(props, 1);                                       /* property count */
  put32(props + 4, PROP_TAG(PT_STRING8, PR_BODY));
  put32(props + 8, 1);                                   /* value count */
  put32(props + 12, 5);                                  /* value size */
  memcpy(props + 16, "hello", 5);                        /* no NUL */
  put32(d, 0x223e9f78); d += 4;                          /* signature */
  *d++ = 1; *d++ = 0;                                    /* key */
  *d++ = 1;                                              /* level */
  put32(d, tag); d += 4;
  put32(d, sizeof(props)); d += 4;
  memcpy(d, props, sizeof(props)); d += sizeof(props);
  for (i = 0; i < sizeof(props); i++) sum += props[i];
  *d++ = sum & 0xff;
  *d++ = (sum >> 8) & 0xff;

  TNEFInitialize(&t);
  if (TNEFParseMemory(file, sizeof(file), &t) < 0) bad++;
  v = MAPIFindProperty(&t.MapiProperties, PROP_TAG(PT_STRING8, PR_BODY));
  if (v == MAPI_UNDEFINED || v->size != 5 || strcmp((char *)v->data, "hello") != 0) bad++;
  TNEFFree(&t);
  return bad;
}

static const struct { const char *name; int (*fn)(void); } cases[] = {
  { "no data", no_data },
  { "shorter than the header", too_short },
  { "uncompressed stream", raw_good },
  { "uncompressed stream claims more than it has", raw_claims_too_much },
  { "compressed stream", compressed_good },
  { "compressed stream: one byte left for a reference", compressed_one_byte_for_reference },
  { "compressed stream: no bytes left for a reference", compressed_no_bytes_for_reference },
  { "compressed stream: no bytes left for a literal", compressed_no_bytes_for_literal },
  { "compressed stream with a huge size", compressed_huge_size },
  { "unknown magic", unknown_magic },
  { "IsCompressedRTF", is_compressed },
  { "parsing stops inside a property", parse_stops_in_property },
  { "string property without its NUL", string_without_nul },
};

int main(void) {
  size_t i;
  int failed = 0;

  for (i = 0; i < sizeof(cases) / sizeof(cases[0]); i++) {
    int status = 0;
    pid_t pid;

    fflush(stdout);
    pid = fork();
    if (pid == 0) {
      /* The library prints its own messages to stdout: keep them out. */
      if (freopen("/dev/null", "w", stdout) == NULL) _exit(2);
      _exit(cases[i].fn());
    }
    waitpid(pid, &status, 0);
    if (WIFEXITED(status) && WEXITSTATUS(status) == 0) {
      printf("PASS: %s\n", cases[i].name);
    } else {
      printf("FAIL: %s\n", cases[i].name);
      failed++;
    }
  }
  return failed ? 1 : 0;
}
