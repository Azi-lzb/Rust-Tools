#include <chrono>
#include <algorithm>
#include <cmath>
#include <ctime>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <sstream>
#include <stdexcept>
#include <string>
#include <vector>
#include <sys/resource.h>

#include "libxl.h"

namespace fs = std::filesystem;
using namespace libxl;
using Clock = std::chrono::steady_clock;

namespace {
double ms(Clock::time_point a, Clock::time_point b) {
    return std::chrono::duration<double, std::milli>(b - a).count();
}
double cpuMs(std::clock_t a, std::clock_t b) {
    return 1000.0 * static_cast<double>(b - a) / CLOCKS_PER_SEC;
}
double peakMemoryMb() {
    struct rusage usage{};
    getrusage(RUSAGE_SELF, &usage);
    return static_cast<double>(usage.ru_maxrss) / 1024.0;
}
std::string csv(const std::string& value) {
    std::string result = "\"";
    for (char c : value) { if (c == '"') result += '"'; result += c; }
    return result + '"';
}
void emit(std::ofstream& out, const std::vector<std::string>& fields) {
    for (size_t i = 0; i < fields.size(); ++i) { if (i) out << ','; out << fields[i]; }
    out << '\n';
}
std::string n(double value) { std::ostringstream s; s << std::fixed << std::setprecision(3) << value; return s.str(); }
std::string sz(const fs::path& p) { return fs::exists(p) ? std::to_string(fs::file_size(p)) : "0"; }
Book* load(const fs::path& path, bool xlsx) {
    Book* book = xlsx ? xlCreateXMLBook() : xlCreateBook();
    if (!book) return nullptr;
    book->setLocale("UTF-8");
    if (!book->load(path.string().c_str())) { book->release(); return nullptr; }
    return book;
}

void runBatch(const fs::path& input, const fs::path& output, int round, std::ofstream& out) {
    const bool xlsx = input.extension() == ".xlsx";
    const auto begin = Clock::now(); const auto cpu0 = std::clock();
    double open = 0, append = 0, edit = 0, del = 0, styleMs = 0, print = 0, save = 0;
    std::string error;
    try {
        auto t = Clock::now();
        Book* book = load(input, xlsx);
        open = ms(t, Clock::now());
        if (!book) throw std::runtime_error("LibXL load failed (possibly evaluation limit)");
        Sheet* sheet = book->getSheet(0);
        if (!sheet) { book->release(); throw std::runtime_error("missing first sheet"); }

        t = Clock::now();
        const int first = sheet->lastRow() + 1;
        for (int i = 0; i < 5; ++i) {
            const std::string label = "perf-" + std::to_string(i);
            if (!sheet->writeStr(first + i, 0, label.c_str()) || !sheet->writeNum(first + i, 1, i))
                throw std::runtime_error("append failed");
        }
        append = ms(t, Clock::now());
        t = Clock::now();
        const std::string editValue = "edit-r" + std::to_string(round);
        // Evaluation builds reserve row 0 for their mandatory banner, matching the phase-1 smoke fixture.
        if (!sheet->writeStr(1, 0, editValue.c_str())) throw std::runtime_error("edit failed");
        edit = ms(t, Clock::now());
        t = Clock::now();
        if (!sheet->insertRow(first + 5, first + 5) || !sheet->removeRow(first + 5, first + 5) ||
            !sheet->insertCol(10, 10) || !sheet->removeCol(10, 10)) throw std::runtime_error("insert/delete failed");
        del = ms(t, Clock::now());
        t = Clock::now();
        Font* font = book->addFont(); font->setBold(true); font->setSize(11);
        Format* format = book->addFormat(); format->setFont(font); format->setNumFormat(NUMFORMAT_NUMBER_D2);
        format->setBorderBottom(BORDERSTYLE_THIN); format->setAlignH(ALIGNH_CENTER);
        if (!sheet->writeNum(first, 1, 0, format) || !sheet->copyCell(first, 1, first, 2)) throw std::runtime_error("style/copy failed");
        styleMs = ms(t, Clock::now());
        t = Clock::now();
        sheet->setPaper(PAPER_A4); sheet->setLandscape(true); sheet->setPrintFit(1, 0);
        sheet->setPrintArea(0, first + 4, 0, 2); sheet->setPrintRepeatRows(0, 0); sheet->setPrintRepeatCols(0, 0);
        sheet->setHeader("Phase 3", 0.2); sheet->setFooter("&P", 0.2);
        print = ms(t, Clock::now());
        t = Clock::now();
        if (!book->save(output.string().c_str())) { std::string e = book->errorMessage() ? book->errorMessage() : "save failed"; book->release(); throw std::runtime_error(e); }
        save = ms(t, Clock::now());
        book->release();

        Book* checkBook = load(output, xlsx);
        if (!checkBook) throw std::runtime_error("output reopen failed");
        Sheet* check = checkBook->getSheet(0);
        const char* raw = check->readStr(first + 4, 0);
        const std::string tail = raw ? raw : "";
        const bool valid = tail == "perf-4" && std::abs(check->readNum(first + 4, 1) - 4.0) < 1e-9;
        checkBook->release();
        if (!valid) throw std::runtime_error("output append verification failed");
        emit(out, {"libxl", std::to_string(round), csv(input.filename().string()), csv(input.extension().string()), sz(input), sz(output),
                   n(open), n(append), n(edit), n(del), n(styleMs), n(print), "0", n(save), "0",
                   n(ms(begin, Clock::now())), n(peakMemoryMb()), n(cpuMs(cpu0, std::clock())), "PASS",
                   csv("calculate=UNSUPPORTED_BY_DESIGN;conversion=RECONSTRUCTION_ONLY")});
    } catch (const std::exception& ex) {
        error = ex.what();
        emit(out, {"libxl", std::to_string(round), csv(input.filename().string()), csv(input.extension().string()), sz(input), sz(output),
                   n(open), n(append), n(edit), n(del), n(styleMs), n(print), "0", n(save), "0",
                   n(ms(begin, Clock::now())), n(peakMemoryMb()), n(cpuMs(cpu0, std::clock())), "FAIL", csv(error)});
    }
}

void runLarge(const fs::path& input, const fs::path& output, std::ofstream& out) {
    const auto begin = Clock::now(); const auto cpu0 = std::clock();
    double open = 0, save = 0; std::string error;
    try {
        auto t = Clock::now(); Book* book = load(input, true); open = ms(t, Clock::now());
        if (!book) throw std::runtime_error("large XLSX load failed");
        Sheet* sheet = book->getSheet(0);
        const int row = sheet->lastFilledRow();
        const double value = sheet->readNum(row, 1);
        if (!sheet->writeStr(1, 0, "large-edited")) { book->release(); throw std::runtime_error("large edit failed"); }
        t = Clock::now(); if (!book->save(output.string().c_str())) { book->release(); throw std::runtime_error("large save failed"); }
        save = ms(t, Clock::now()); book->release();
        emit(out, {"libxl", "1", csv(input.filename().string()), csv("xlsx-large"), sz(input), sz(output), n(open), "0", "0", "0", "0", "0", "0", n(save), "0",
                   n(ms(begin, Clock::now())), n(peakMemoryMb()), n(cpuMs(cpu0, std::clock())), "PASS", csv("lastRow=" + std::to_string(row + 1) + ";lastB=" + n(value))});
    } catch (const std::exception& ex) {
        error = ex.what();
        emit(out, {"libxl", "1", csv(input.filename().string()), csv("xlsx-large"), sz(input), sz(output), n(open), "0", "0", "0", "0", "0", "0", n(save), "0",
                   n(ms(begin, Clock::now())), n(peakMemoryMb()), n(cpuMs(cpu0, std::clock())), "FAIL", csv(error)});
    }
}
}  // namespace

int main(int argc, char** argv) {
    if (argc != 4) { std::cerr << "usage: libxl_perf INPUT_DIRECTORY OUTPUT_DIRECTORY CSV_PATH\n"; return 2; }
    const fs::path input(argv[1]), output(argv[2]); fs::create_directories(output);
    std::ofstream csvFile(argv[3]);
    csvFile << "engine,round,file,format,input_bytes,output_bytes,open_ms,append_ms,edit_ms,delete_rows_cols_ms,style_ms,print_setup_ms,calculate_ms,save_ms,convert_ms,total_ms,peak_memory_mb,cpu_ms,status,error\n";
    std::vector<fs::path> batch;
    for (const auto& item : fs::directory_iterator(input)) if (item.is_regular_file() && item.path().filename().string().rfind("batch-", 0) == 0) batch.push_back(item.path());
    std::sort(batch.begin(), batch.end());
    if (!batch.empty()) {
        Book* warm = load(batch.front(), batch.front().extension() == ".xlsx");
        if (warm) { warm->save((output / "warmup.xls").string().c_str()); warm->release(); }
        std::cout << "WARMUP_COMPLETE\n";
    }
    for (int round = 1; round <= 3; ++round)
        for (const auto& file : batch) runBatch(file, output / ("r" + std::to_string(round) + "-" + file.filename().string()), round, csvFile);
    runLarge(input / "large-150k.xlsx", output / "large-150k.xlsx", csvFile);
    runLarge(input / "large-1m.xlsx", output / "large-1m.xlsx", csvFile);
    csvFile.flush();
    return 0;
}
