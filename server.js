const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = 3000;

app.use(express.static("public"));

if (!fs.existsSync("uploads")) fs.mkdirSync("uploads");
if (!fs.existsSync("output")) fs.mkdirSync("output");

const upload = multer({ dest: "uploads/" });

/* ================= BUSINESS CONFIG ================= */

const WORK_START_HOUR = 9;
const WORK_START_MIN = 15;
const WORK_END_HOUR = 18;
const STANDARD_DAILY_HOURS = 8;

/* =================================================== */

function formatTime(dateObj) {
    return dateObj.toTimeString().slice(0, 5);
}

function parseExcelTime(value) {
    if (value instanceof Date) return value;
    if (typeof value === "number") {
        const t = XLSX.SSF.parse_date_code(value);
        return new Date(0, 0, 0, t.H, t.M, t.S);
    }
    return new Date(`1970-01-01T${value}`);
}

function hoursDiff(start, end) {
    return (end - start) / (1000 * 60 * 60);
}

app.post("/upload", upload.single("file"), (req, res) => {
    try {

        const workbook = XLSX.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        let data = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        let employees = {};

        /* ================= READ EXCEL ================= */

        data.forEach(row => {

            const { Id, Name, Date: ExcelDate, Time, Status } = row;
            if (!Id || !Name || !ExcelDate || !Time || !Status) return;

            const dateObj = new Date(ExcelDate);
            const timeObj = parseExcelTime(Time);

            const fullDateTime = new Date(
                dateObj.getFullYear(),
                dateObj.getMonth(),
                dateObj.getDate(),
                timeObj.getHours(),
                timeObj.getMinutes(),
                timeObj.getSeconds()
            );

            if (!employees[Id]) {
                employees[Id] = {
                    id: Id,
                    name: Name,
                    punches: []
                };
            }

            employees[Id].punches.push({
                datetime: fullDateTime,
                status: Status.trim().toLowerCase()
            });
        });

        let attendanceSheet = [];
        let summarySheet = [];

        /* ================= PROCESS EACH EMPLOYEE ================= */

        Object.values(employees)
            .sort((a, b) => a.id - b.id)
            .forEach(emp => {

                if (emp.punches.length === 0) return;

                /* Remove duplicates */
                emp.punches = emp.punches.filter((v, i, arr) =>
                    arr.findIndex(x =>
                        x.datetime.getTime() === v.datetime.getTime() &&
                        x.status === v.status
                    ) === i
                );

                emp.punches.sort((a, b) => a.datetime - b.datetime);

                const month = emp.punches[0].datetime.getMonth();
                const year = emp.punches[0].datetime.getFullYear();
                const daysInMonth = new Date(year, month + 1, 0).getDate();

                let dailyMap = {};
                for (let d = 1; d <= daysInMonth; d++) {
                    dailyMap[d] = [];
                }

                /* ================= PAIRING ENGINE ================= */

                for (let i = 0; i < emp.punches.length; i++) {

                    const punch = emp.punches[i];
                    if (punch.status !== "check in") continue;

                    let checkIn = punch.datetime;
                    let checkOut = null;

                    for (let j = i + 1; j < emp.punches.length; j++) {
                        if (emp.punches[j].status === "check out") {
                            checkOut = emp.punches[j].datetime;
                            i = j;
                            break;
                        }
                    }

                    if (!checkOut) continue;

                    let targetDay;

                    // Same day shift
                    if (
                        checkIn.getFullYear() === checkOut.getFullYear() &&
                        checkIn.getMonth() === checkOut.getMonth() &&
                        checkIn.getDate() === checkOut.getDate()
                    ) {
                        targetDay = checkIn.getDate();
                    }
                    // Night shift → assign to checkout date
                    else {
                        targetDay = checkOut.getDate();
                    }

                    if (!dailyMap[targetDay]) {
                        dailyMap[targetDay] = [];
                    }

                    dailyMap[targetDay].push({ checkIn, checkOut });
                }

                /* ================= BUILD ATTENDANCE ROW ================= */

                let attendanceRow = { ID: emp.id, Name: emp.name };

                let present = 0, late = 0, early = 0;
                let totalHours = 0, otHours = 0;

                for (let d = 1; d <= daysInMonth; d++) {

                    let shifts = dailyMap[d];

                    if (shifts.length === 0) {
                        attendanceRow[d] = "";
                        continue;
                    }

                    present++;

                    let shiftStrings = [];

                    shifts.forEach(s => {

                        const worked = hoursDiff(s.checkIn, s.checkOut);
                        totalHours += worked;

                        if (worked > STANDARD_DAILY_HOURS)
                            otHours += worked - STANDARD_DAILY_HOURS;

                        if (
                            s.checkIn.getHours() > WORK_START_HOUR ||
                            (s.checkIn.getHours() === WORK_START_HOUR &&
                             s.checkIn.getMinutes() > WORK_START_MIN)
                        ) late++;

                        if (s.checkOut.getHours() < WORK_END_HOUR)
                            early++;

                        shiftStrings.push(
                            `${formatTime(s.checkIn)}-${formatTime(s.checkOut)}`
                        );
                    });

                    attendanceRow[d] = shiftStrings.join(" / ");
                }

                attendanceSheet.push(attendanceRow);

                summarySheet.push({
                    ID: emp.id,
                    Name: emp.name,
                    Present: present,
                    Absent: daysInMonth - present,
                    Late: late,
                    Early: early,
                    "Total Hours": totalHours.toFixed(2),
                    "OT Hours": otHours.toFixed(2)
                });
            });

        /* ================= CREATE EXCEL ================= */

        const wb = XLSX.utils.book_new();

        // Proper header order
        let maxDays = 31;
        if (attendanceSheet.length > 0) {
            maxDays = Object.keys(attendanceSheet[0])
                .filter(k => !isNaN(k)).length;
        }

        let headers = ["ID", "Name"];
        for (let d = 1; d <= maxDays; d++) {
            headers.push(String(d));
        }

        const attWS = XLSX.utils.json_to_sheet(attendanceSheet, {
            header: headers
        });

        const sumWS = XLSX.utils.json_to_sheet(summarySheet);

        XLSX.utils.book_append_sheet(wb, attWS, "Attendance");
        XLSX.utils.book_append_sheet(wb, sumWS, "Summary");

        const outputFile = path.join("output", "Professional_Attendance_Report.xlsx");
        XLSX.writeFile(wb, outputFile);

        fs.unlinkSync(req.file.path);

        res.download(outputFile);

    } catch (err) {
        console.error("FULL ERROR:", err);
        res.status(500).send(err.message);
    }
});

app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
});