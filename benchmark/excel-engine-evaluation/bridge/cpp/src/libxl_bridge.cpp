#include "libxl_bridge.h"

#include <algorithm>
#include <memory>
#include <string>
#include <vector>

#include "libxl.h"

namespace {
struct Workbook {
    libxl::Book* book{};
    libxl::Sheet* sheet{};
};
thread_local std::string error_text;

int fail(const char* message) {
    error_text = message ? message : "unknown LibXL error";
    return 1;
}
}

extern "C" xl_handle xl_open(const char* input_utf8, int xlsx) {
    try {
        if (!input_utf8) { fail("input path is null"); return nullptr; }
        auto holder = std::make_unique<Workbook>();
        holder->book = xlsx ? xlCreateXMLBook() : xlCreateBook();
        if (!holder->book) { fail("LibXL book creation failed"); return nullptr; }
        holder->book->setLocale("UTF-8");
        if (!holder->book->load(input_utf8)) {
            const char* raw = holder->book->errorMessage();
            error_text = raw ? raw : "LibXL load failed";
            holder->book->release(); holder->book = nullptr;
            return nullptr;
        }
        holder->sheet = holder->book->getSheet(0);
        if (!holder->sheet) {
            holder->book->release(); holder->book = nullptr;
            fail("workbook has no first worksheet"); return nullptr;
        }
        error_text.clear();
        return holder.release();
    } catch (const std::exception& e) { fail(e.what()); return nullptr; }
      catch (...) { fail("unknown C++ exception in xl_open"); return nullptr; }
}

extern "C" int xl_append_rows(xl_handle opaque, const char* const* labels,
                                const double* values, size_t count) {
    try {
        auto* holder = static_cast<Workbook*>(opaque);
        if (!holder || !holder->sheet || (count && (!labels || !values))) return fail("invalid append arguments");
        // LibXL's lastRow() reports the used row count (exclusive end), while
        // writeStr/writeNum take zero-based row indices.
        int row = std::max(0, holder->sheet->lastRow());
        for (size_t i = 0; i < count; ++i, ++row) {
            if (!labels[i] || !holder->sheet->writeStr(row, 0, labels[i]) ||
                !holder->sheet->writeNum(row, 1, values[i])) {
                const char* raw = holder->book->errorMessage();
                error_text = raw ? raw : "LibXL append failed"; return 1;
            }
        }
        error_text.clear(); return 0;
    } catch (const std::exception& e) { return fail(e.what()); }
      catch (...) { return fail("unknown C++ exception in xl_append_rows"); }
}

extern "C" int xl_save(xl_handle opaque, const char* output_utf8) {
    try {
        auto* holder = static_cast<Workbook*>(opaque);
        if (!holder || !holder->book || !output_utf8) return fail("invalid save arguments");
        if (!holder->book->save(output_utf8)) {
            const char* raw = holder->book->errorMessage();
            error_text = raw ? raw : "LibXL save failed"; return 1;
        }
        error_text.clear(); return 0;
    } catch (const std::exception& e) { return fail(e.what()); }
      catch (...) { return fail("unknown C++ exception in xl_save"); }
}

extern "C" int xl_verify_last_row(xl_handle opaque, const char* expected_label, double expected_value) {
    try {
        auto* holder = static_cast<Workbook*>(opaque);
        if (!holder || !holder->sheet || !expected_label) return fail("invalid verification arguments");
        const int row = std::max(0, holder->sheet->lastRow() - 1);
        const char* raw = holder->sheet->readStr(row, 0);
        const double actual = holder->sheet->readNum(row, 1);
        if (!raw || std::string(raw) != expected_label || actual != expected_value) {
            error_text = "saved last row mismatch at row=" + std::to_string(row) +
                " label=" + (raw ? std::string(raw) : std::string("<null>")) +
                " expected=" + expected_label + " number=" + std::to_string(actual) +
                " expected_number=" + std::to_string(expected_value);
            return 1;
        }
        error_text.clear(); return 0;
    } catch (const std::exception& e) { return fail(e.what()); }
      catch (...) { return fail("unknown C++ exception in xl_verify_last_row"); }
}

extern "C" void xl_close(xl_handle opaque) {
    auto* holder = static_cast<Workbook*>(opaque);
    if (!holder) return;
    if (holder->book) holder->book->release();
    delete holder;
}

extern "C" const char* xl_last_error(void) { return error_text.c_str(); }
