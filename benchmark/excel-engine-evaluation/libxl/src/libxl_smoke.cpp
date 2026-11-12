#include <cmath>
#include <filesystem>
#include <iostream>
#include <string>

#include "libxl.h"

namespace fs = std::filesystem;
using namespace libxl;

namespace {
constexpr int kTextRow = 1;
constexpr int kNumberRow = 2;
constexpr int kSecondNumberRow = 3;
constexpr int kFormulaRow = 4;

void fail(const std::string& message) {
    std::cerr << "FAIL " << message << '\n';
    std::exit(1);
}

void createFixture(const fs::path& path, bool xlsx) {
    Book* book = xlsx ? xlCreateXMLBook() : xlCreateBook();
    if (!book) fail("create workbook");
    book->setLocale("UTF-8");
    Sheet* sheet = book->addSheet("Data");
    if (!sheet) {
        book->release();
        fail("add Data sheet");
    }

    // The evaluation build reserves the first row for its banner.
    const bool wrote = sheet->writeStr(kTextRow, 0, "中文读取验证") &&
                       sheet->writeNum(kNumberRow, 0, 10.0) &&
                       sheet->writeNum(kSecondNumberRow, 0, 20.0) &&
                       sheet->writeFormula(kFormulaRow, 0, "=SUM(A3:A4)");
    if (!wrote || !book->save(path.string().c_str())) {
        const std::string error = book->errorMessage() ? book->errorMessage() : "unknown";
        book->release();
        fail("create fixture " + path.string() + ": " + error);
    }
    book->release();
}

Book* loadBook(const fs::path& path, bool xlsx) {
    Book* book = xlsx ? xlCreateXMLBook() : xlCreateBook();
    if (!book) fail("create reader for " + path.string());
    book->setLocale("UTF-8");
    if (!book->load(path.string().c_str())) {
        const std::string error = book->errorMessage() ? book->errorMessage() : "unknown";
        book->release();
        fail("load " + path.string() + ": " + error);
    }
    return book;
}

void verifyFixture(const fs::path& path, bool xlsx,
                   const std::string& expectedText, double expectedNumber,
                   bool requireFormula = true) {
    Book* book = loadBook(path, xlsx);
    Sheet* sheet = book->getSheet(0);
    if (!sheet) {
        book->release();
        fail("read first worksheet from " + path.string());
    }
    const std::string sheetName = book->getSheetName(0) ? book->getSheetName(0) : "";
    const char* textRaw = sheet->readStr(kTextRow, 0);
    const std::string text = textRaw ? textRaw : "";
    const double number = sheet->readNum(kNumberRow, 0);
    const char* formulaRaw = sheet->readFormula(kFormulaRow, 0);
    const std::string formula = formulaRaw ? formulaRaw : "";
    const bool ok = book->sheetCount() >= 1 &&
                    sheetName == "Data" && text == expectedText &&
                    std::abs(number - expectedNumber) < 1e-9 &&
                    (!requireFormula || !formula.empty());
    std::cout << "READ file=" << path.filename().string()
              << " sheet=" << sheetName
              << " text=" << text
              << " number=" << number
              << " formula=" << formula << '\n';
    book->release();
    if (!ok) fail("readback values mismatch for " + path.string());
}

void modifyFixture(const fs::path& input, const fs::path& output, bool xlsx) {
    Book* book = loadBook(input, xlsx);
    Sheet* sheet = book->getSheet(0);
    if (!sheet || !sheet->writeNum(kNumberRow, 0, 30.0) ||
        !sheet->writeStr(kTextRow, 0, "中文已修改") ||
        !book->save(output.string().c_str())) {
        const std::string error = book->errorMessage() ? book->errorMessage() : "unknown";
        book->release();
        fail("modify/save " + output.string() + ": " + error);
    }
    book->release();
    verifyFixture(output, xlsx, "中文已修改", 30.0);
    std::cout << "MODIFY format=" << (xlsx ? "xlsx" : "xls") << " PASS\n";
}

void tryDirectXlsToXlsx(const fs::path& input, const fs::path& output) {
    Book* book = xlCreateXMLBook();
    if (!book) fail("create xlsx converter");
    book->setLocale("UTF-8");
    if (!book->load(input.string().c_str())) {
        const std::string error = book->errorMessage() ? book->errorMessage() : "unknown";
        std::cout << "XLS_TO_XLSX_DIRECT UNSUPPORTED load=" << error << '\n';
        book->release();
        return;
    }
    if (!book->save(output.string().c_str())) {
        const std::string error = book->errorMessage() ? book->errorMessage() : "unknown";
        std::cout << "XLS_TO_XLSX_DIRECT UNSUPPORTED save=" << error << '\n';
        book->release();
        return;
    }
    book->release();
    verifyFixture(output, true, "中文读取验证", 10.0, false);
    std::cout << "XLS_TO_XLSX_DIRECT PASS\n";
}

void convertSimpleXlsToXlsx(const fs::path& input, const fs::path& output) {
    Book* sourceBook = loadBook(input, false);
    Sheet* sourceSheet = sourceBook->getSheet(0);
    Book* outputBook = xlCreateXMLBook();
    if (!sourceSheet || !outputBook) {
        sourceBook->release();
        if (outputBook) outputBook->release();
        fail("create simple XLS to XLSX transcode");
    }
    outputBook->setLocale("UTF-8");
    const char* sourceSheetName = sourceBook->getSheetName(0);
    Sheet* outputSheet = outputBook->addSheet(sourceSheetName ? sourceSheetName : "Data");
    if (!outputSheet) {
        sourceBook->release();
        outputBook->release();
        fail("add output sheet for XLS to XLSX transcode");
    }

    bool copied = true;
    for (int row = kTextRow; row <= kFormulaRow && copied; ++row) {
        const char* formulaRaw = sourceSheet->readFormula(row, 0);
        const std::string formula = formulaRaw ? formulaRaw : "";
        if (!formula.empty()) {
            const std::string expression = formula.front() == '=' ? formula : "=" + formula;
            copied = outputSheet->writeFormula(row, 0, expression.c_str());
            continue;
        }
        switch (sourceSheet->cellType(row, 0)) {
            case CELLTYPE_STRING: {
                const char* valueRaw = sourceSheet->readStr(row, 0);
                const std::string value = valueRaw ? valueRaw : "";
                copied = outputSheet->writeStr(row, 0, value.c_str());
                break;
            }
            case CELLTYPE_NUMBER:
                copied = outputSheet->writeNum(row, 0, sourceSheet->readNum(row, 0));
                break;
            case CELLTYPE_BOOLEAN:
                copied = outputSheet->writeBool(row, 0, sourceSheet->readBool(row, 0));
                break;
            default:
                break;
        }
    }

    const bool saved = copied && outputBook->save(output.string().c_str());
    const std::string sourceError = sourceBook->errorMessage() ? sourceBook->errorMessage() : "unknown";
    const std::string outputError = outputBook->errorMessage() ? outputBook->errorMessage() : "unknown";
    sourceBook->release();
    outputBook->release();
    if (!saved) fail("save simple XLS to XLSX transcode: " + sourceError + "; " + outputError);
    verifyFixture(output, true, "中文读取验证", 10.0, false);
    std::cout << "XLS_TO_XLSX_SAMPLE_COPY PASS (scalar values and formula only)\n";
}

void capabilityTest(const fs::path& input, const fs::path& output, bool xlsx) {
    Book* book = loadBook(input, xlsx);
    Sheet* sheet = book->getSheet(0);
    if (!sheet) { book->release(); fail("capability test missing worksheet"); }

    // The evaluation license caps accessible cells at 300. Exercise a bounded bulk append (100 x 2).
    for (int i = 0; i < 100; ++i) {
        if (!sheet->writeStr(5 + i, 0, ("append-" + std::to_string(i)).c_str()) ||
            !sheet->writeNum(5 + i, 1, static_cast<double>(i))) {
            const std::string error = book->errorMessage() ? book->errorMessage() : "unknown";
            book->release(); fail("append rows: " + error);
        }
    }
    if (!sheet->insertRow(105, 105) || !sheet->removeRow(105, 105) ||
        !sheet->insertCol(2, 2) || !sheet->removeCol(2, 2)) {
        const std::string error = book->errorMessage() ? book->errorMessage() : "unknown";
        book->release(); fail("insert/delete rows/columns: " + error);
    }

    Font* font = book->addFont();
    font->setBold(true); font->setSize(12); font->setColor(COLOR_DARKBLUE);
    Format* format = book->addFormat();
    format->setFont(font);
    format->setNumFormat(NUMFORMAT_NUMBER_D2);
    format->setFillPattern(FILLPATTERN_SOLID);
    format->setPatternForegroundColor(COLOR_LIGHTYELLOW);
    format->setBorderBottom(BORDERSTYLE_THIN);
    format->setBorderBottomColor(COLOR_BLUE);
    format->setAlignH(ALIGNH_CENTER);
    format->setWrap(true);
    if (!sheet->writeNum(5, 1, 0.0, format) || !sheet->copyCell(5, 1, 5, 2)) {
        const std::string error = book->errorMessage() ? book->errorMessage() : "unknown";
        book->release(); fail("style/copy style: " + error);
    }
    if (!sheet->setCol(0, 1, 22.0) || !sheet->setRow(5, 24.0) ||
        !sheet->setColHidden(8, true) || !sheet->setColHidden(8, false) ||
        !sheet->setRowHidden(3, true) || !sheet->setRowHidden(3, false)) {
        const std::string error = book->errorMessage() ? book->errorMessage() : "unknown";
        book->release(); fail("dimensions/visibility: " + error);
    }
    sheet->setPaper(PAPER_A4);
    sheet->setLandscape(true);
    sheet->setMarginLeft(0.4); sheet->setMarginRight(0.4);
    sheet->setMarginTop(0.5); sheet->setMarginBottom(0.5);
    sheet->setPrintFit(1, 0);
    sheet->setHCenter(true); sheet->setVCenter(false);
    sheet->setHeader("&BPhase 2", 0.2);
    sheet->setFooter("Page &P of &N", 0.2);
    sheet->setPrintArea(1, 104, 0, 2);
    sheet->setPrintRepeatRows(0, 0);
    sheet->setPrintRepeatCols(0, 0);
    if (!book->save(output.string().c_str())) {
        const std::string error = book->errorMessage() ? book->errorMessage() : "unknown";
        book->release(); fail("save capability workbook: " + error);
    }
    book->release();

    Book* checkBook = loadBook(output, xlsx);
    Sheet* check = checkBook->getSheet(0);
    const char* appendedRaw = check->readStr(104, 0);
    const std::string appended = appendedRaw ? appendedRaw : "";
    const double appendedNumber = check->readNum(104, 1);
    Format* savedFormat = check->cellFormat(5, 1);
    int rowFirst = -1, rowLast = -1, colFirst = -1, colLast = -1;
    int repeatedFirst = -1, repeatedLast = -1;
    int fitWide = -1, fitTall = -1;
    const bool areaOk = check->printArea(&rowFirst, &rowLast, &colFirst, &colLast) &&
                        rowFirst == 1 && rowLast == 104 && colFirst == 0 && colLast == 2;
    const bool repeatOk = check->printRepeatRows(&repeatedFirst, &repeatedLast) &&
                          repeatedFirst == 0 && repeatedLast == 0;
    const bool fitOk = check->getPrintFit(&fitWide, &fitTall) && fitWide == 1 && fitTall == 0;
    const bool ok = appended == "append-99" && std::abs(appendedNumber - 99.0) < 1e-9 &&
                    savedFormat && savedFormat->font() && savedFormat->font()->bold() &&
                    std::abs(check->colWidth(0) - 22.0) < 0.1 &&
                    std::abs(check->rowHeight(5) - 24.0) < 0.1 &&
                    check->paper() == PAPER_A4 && check->landscape() &&
                    check->hCenter() && !check->vCenter() && areaOk && repeatOk && fitOk &&
                    check->header() && check->footer();
    std::cout << "CAPABILITY format=" << (xlsx ? "xlsx" : "xls")
              << " append=" << appended << "," << appendedNumber
              << " width=" << check->colWidth(0) << " height=" << check->rowHeight(5)
              << " printArea=" << rowFirst << ":" << rowLast << "," << colFirst << ":" << colLast
              << " fit=" << fitWide << ":" << fitTall << " check=" << (ok ? "PASS" : "FAIL") << '\n';
    checkBook->release();
    if (!ok) fail("capability save/reopen assertions failed for " + output.string());
}
}  // namespace

int main(int argc, char** argv) {
    if (argc != 2) {
        std::cerr << "usage: libxl_smoke OUTPUT_DIRECTORY\n";
        return 2;
    }
    const fs::path out(argv[1]);
    fs::create_directories(out);
    const fs::path xls = out / "libxl-input.xls";
    const fs::path xlsx = out / "libxl-input.xlsx";
    const fs::path modifiedXls = out / "libxl-modified.xls";
    const fs::path modifiedXlsx = out / "libxl-modified.xlsx";
    const fs::path convertedXlsx = out / "libxl-converted.xlsx";

    createFixture(xls, false);
    createFixture(xlsx, true);
    verifyFixture(xls, false, "中文读取验证", 10.0);
    verifyFixture(xlsx, true, "中文读取验证", 10.0);
    modifyFixture(xls, modifiedXls, false);
    modifyFixture(xlsx, modifiedXlsx, true);
    tryDirectXlsToXlsx(xls, convertedXlsx);
    convertSimpleXlsToXlsx(xls, convertedXlsx);
    capabilityTest(xls, out / "libxl-capabilities.xls", false);
    capabilityTest(xlsx, out / "libxl-capabilities.xlsx", true);

    std::cout << "FORMULA_RECALC UNSUPPORTED (LibXL has no calculation engine)\n";
    std::cout << "SMOKE_COMPLETE\n";
    return 0;
}
