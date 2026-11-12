#include <cmath>
#include <cstddef>
#include <cstdlib>
#include <iostream>
#include <string>

// The 16.7.0 Linux headers conditionally define their global `byte` alias
// based on this C++17 feature macro; undefine it after <cstddef> so the
// vendor's own typedef is visible before its umbrella header is parsed.
#ifdef __cpp_lib_byte
#undef __cpp_lib_byte
#endif
#include "Spire.Xls.o.h"

using Spire::Xls::Workbook;

namespace {
void fail(const std::string& message) {
    std::cerr << "FAIL " << message << '\n';
    std::exit(1);
}

void smoke(const wchar_t* input, const wchar_t* output, const char* label) {
    Workbook workbook;
    workbook.LoadFromFile(input);
    auto worksheets = workbook.GetWorksheets();
    if (worksheets->GetCount() < 1) fail(std::string(label) + " no worksheets");
    auto sheet = worksheets->Get(0);
    auto cell = sheet->GetRange()->Get(3, 1);
    const double before = cell->GetNumberValue();
    if (!std::isfinite(before) || std::fabs(before - 10.0) > 1e-9) {
        fail(std::string(label) + " expected the sample value 10 in A3");
    }
    const double after = before + 1.0;
    cell->SetNumberValue(after);
    workbook.SaveToFile(output);

    Workbook reopened;
    reopened.LoadFromFile(output);
    auto check = reopened.GetWorksheets()->Get(0)->GetRange()->Get(3, 1);
    if (std::fabs(check->GetNumberValue() - after) > 1e-9) {
        fail(std::string(label) + " saved value did not round-trip");
    }
    std::cout << "PASS " << label << " read=" << before << " saved=" << after << '\n';
}
}  // namespace

int main() {
    try {
        smoke(L"/output/input.xlsx", L"/output/modified.xlsx", "XLSX");
        smoke(L"/output/input.xls", L"/output/modified.xls", "XLS");
    } catch (const std::exception& ex) {
        fail(ex.what());
    } catch (...) {
        fail("unknown Spire exception");
    }
    return 0;
}
