#include <cmath>
#include <cstdlib>
#include <iostream>
#include <string>

#include <Aspose.Cells.h>

using Aspose::Cells::SaveFormat;
using Aspose::Cells::Workbook;

namespace {
void fail(const std::string& message) {
    std::cerr << "FAIL " << message << '\n';
    std::exit(1);
}

void smoke(const char16_t* input, const char16_t* output, SaveFormat format,
           const char* label) {
    Workbook workbook(input);
    auto sheets = workbook.GetWorksheets();
    if (sheets.GetCount() < 1) fail(std::string(label) + " no worksheets");
    auto sheet = sheets.Get(0);
    auto cells = sheet.GetCells();
    auto cell = cells.Get(2, 0);  // Aspose indexes are zero-based: A3.
    const double before = cell.GetDoubleValue();
    if (!std::isfinite(before) || std::fabs(before - 10.0) > 1e-9) {
        fail(std::string(label) + " expected the sample value 10 in A3");
    }
    cell.PutValue(before + 1.0);
    workbook.Save(output, format);

    Workbook reopened(output);
    const double after = reopened.GetWorksheets().Get(0).GetCells().Get(2, 0).GetDoubleValue();
    if (std::fabs(after - 11.0) > 1e-9) fail(std::string(label) + " did not round-trip");
    std::cout << "PASS " << label << " read=" << before << " saved=" << after << '\n';
}
}  // namespace

int main() {
    try {
        smoke(u"/output/input.xlsx", u"/output/modified.xlsx", SaveFormat::Xlsx, "XLSX");
        smoke(u"/output/input.xls", u"/output/modified.xls", SaveFormat::Excel97To2003, "XLS");
    } catch (const std::exception& ex) {
        fail(ex.what());
    } catch (...) {
        fail("unknown Aspose.Cells C++ exception");
    }
    return 0;
}
