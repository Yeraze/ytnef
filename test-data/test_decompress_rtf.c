/*
 * DecompressRTF() must refuse a property with no usable data rather than
 * reading through its data pointer.
 */
#include <stdio.h>
#include <ytnef.h>

static int check(const char *name, BYTE *data, int size) {
  variableLength p;
  int out = 0;

  p.data = data;
  p.size = size;
  if (DecompressRTF(&p, &out) != NULL) {
    printf("FAIL: %s: expected NULL\n", name);
    return 1;
  }
  printf("PASS: %s\n", name);
  return 0;
}

int main(void) {
  BYTE shortbuf[4] = { 0, 0, 0, 0 };
  int failed = 0;

  failed += check("no data, no size", NULL, 0);
  failed += check("no data, header-sized", NULL, 20);
  failed += check("no data, large size", NULL, 0x2b0000);
  failed += check("data shorter than the header", shortbuf, sizeof(shortbuf));
  return failed ? 1 : 0;
}
