#pragma once

#include <stddef.h>

#ifdef __cplusplus
extern "C" {
#endif

typedef void* xl_handle;
xl_handle xl_open(const char* input_utf8, int xlsx);
int xl_append_rows(xl_handle handle, const char* const* labels, const double* values, size_t count);
int xl_save(xl_handle handle, const char* output_utf8);
int xl_verify_last_row(xl_handle handle, const char* expected_label, double expected_value);
void xl_close(xl_handle handle);
const char* xl_last_error(void);

#ifdef __cplusplus
}
#endif
