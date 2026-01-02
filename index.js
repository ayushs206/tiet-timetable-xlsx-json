const xlsx = require('xlsx');
const fs = require('fs');

const wb = xlsx.readFile("timetable.xlsx");

let subjectMap = {
    "UPH013": "Applied Physics",
    "UMA023": "Linear Algebra",
    "UES101": "Engineering Drawing",
    "UHU003": "Proffessional Communication",
    "UES102": "Manufacturing Processes"
}

const batchGroups = {
    "1A1": ["1A11", "1A12", "1A13", "1A14", "1A15", "1A16", "1A17", "1A18"],
    "1A2": ["1A21", "1A22", "1A23", "1A24", "1A25", "1A26", "1A27", "1A28"],
    "1A3": ["1A31", "1A32", "1A33", "1A34", "1A35", "1A36", "1A37", "1A38"],
    "1A4": ["1A41", "1A42", "1A43", "1A44", "1A45"],
    "1A5": ["1A51", "1A52", "1A53", "1A54", "1A55"],
    "1A6": ["1A61", "1A62", "1A63", "1A64", "1A65"],
    "1A7": ["1A71", "1A72", "1A73", "1A74", "1A75"],
    "1A8": ["1A81", "1A82", "1A83", "1A84", "1A85"],
    "1A9": ["1A91", "1A92", "1A93", "1A94", "1A95"]
};


let eventMap = {
    "L": "Lecture",
    "P": "Practical",
    "T": "Tutorial"
}

const sheet = wb.Sheets[wb.SheetNames[0]];
// const grid = xlsx.utils.sheet_to_json(sheet);
const grid = JSON.parse(
    fs.readFileSync("grid.json", "utf8")
);

// fs.writeFileSync("grid.json", JSON.stringify(grid, null, 2));

let batchesIndexed = {};
let jsonbatched = Object.entries(grid[3]);

for (const jsonbatch of jsonbatched) {
    const startindx = jsonbatch[0].split('_')[3]
    batchesIndexed[jsonbatch[1]] = { indx: parseInt(startindx) }
    if (jsonbatch[1] === "1A95") break;
}

let grid5 = Object.entries(grid[5]);
for (const grid of grid5) {
    let gridIndx = grid[0].split('_')[3];
    if (parseInt(gridIndx) < 4) continue;
}

function normalizeTime(val) {
    let h, m, ap;

    if (typeof val === "number") {
        // Excel time fraction
        const mins = Math.round(val * 24 * 60);
        h = Math.floor(mins / 60);
        m = mins % 60;
        ap = h >= 12 ? "PM" : "AM";
        h = h % 12 || 12;
    } else {
        const s = String(val).replace(/\s+/g, " ").trim();
        const match = s.match(/^(\d{1,2}):(\d{1,2})\s*(AM|PM)$/i);
        if (!match) return s;

        h = parseInt(match[1], 10);
        m = parseInt(match[2], 10);
        ap = match[3].toUpperCase();
    }

    return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")} ${ap}`;
}

let currentDay = null;
const dayByMgr = {
    5: "Monday",
    33: "Tuesday",
    61: "Wednesday",
    89: "Thursday",
    117: "Friday",
    131: "Saturday"
};

// let newdata = {}
// for (i = 0; i < grid.length; i++) {
//     if (!grid[i]["__EMPTY"]) continue;
//     newdata[i] = grid[i]["__EMPTY"]
// }

// fs.writeFileSync("result2.json", JSON.stringify(newdata, null, 2));
// return;

for (let mgr = 5; mgr <= grid.length - 2; mgr += 2) {

    const row = grid[mgr]
    const row2 = grid[mgr + 1]

    if (dayByMgr[mgr]) {
        currentDay = dayByMgr[mgr];
    }
    if (!currentDay) continue;

    const day = currentDay;
    if (!day) continue;

    const timeCell = row["__EMPTY_3"];
    if (!timeCell) continue;
    const time = normalizeTime(timeCell);
    if (time === "05:10 PM" || time === "06:50 PM") continue;

    for (i = 4; i <= 110; i += 2) {

        const key = `__EMPTY_${i}`
        if (!(key in row)) continue;

        const value = row[key];
        let subjectCode = splitLastChar(value)
        if (!subjectMap[subjectCode.main]) continue;

        const batch = getBatchFromIndex(i);

        if (subjectCode.last === "L" || (subjectCode.last === "P" && subjectCode.main === "UES102")) {

            const group = getGroupFromBatch(batch);

            const targets = batchGroups[group];
            if (!targets) continue;

            let newTime = getNextSlot(time);

            for (const b of targets) {
                batchesIndexed[b][day] = batchesIndexed[b][day] || {};
                batchesIndexed[b][day][time] = [
                    subjectCode.main,
                    row2[key],
                    subjectMap[subjectCode.main],
                    eventMap[subjectCode.last]
                ];
                if (subjectCode.last === "P") {
                    batchesIndexed[b][day][newTime] = [
                        subjectCode.main,
                        row2[key],
                        subjectMap[subjectCode.main],
                        eventMap[subjectCode.last]
                    ]
                }
            }

            continue; // IMPORTANT: prevent double write
        }

        batchesIndexed[batch][day] = batchesIndexed[batch][day] || {}

        batchesIndexed[batch][day][time] = [
            subjectCode.main,
            row2[key],
            subjectMap[subjectCode.main],
            eventMap[subjectCode.last]
        ];


        if (subjectCode.last === "P" || (subjectCode.last === "T" && subjectCode.main === "UES101")) {
            let newTime = getNextSlot(time);
            batchesIndexed[batch][day] = batchesIndexed[batch][day] || {}

            batchesIndexed[batch][day][newTime] = [
                subjectCode.main,
                row2[key],
                subjectMap[subjectCode.main],
                eventMap[subjectCode.last]
            ];
        }
    }
}

fs.writeFileSync("result.json", JSON.stringify(batchesIndexed, null, 2));

function getNextSlot(time) {
    const match = time.match(/^(\d{2}):(\d{2}) (AM|PM)$/);
    if (!match) return null;

    let h = parseInt(match[1], 10);
    let m = parseInt(match[2], 10);
    let ap = match[3];

    // convert to minutes since midnight
    if (ap === "PM" && h !== 12) h += 12;
    if (ap === "AM" && h === 12) h = 0;

    let total = h * 60 + m + 50; // add 50 minutes

    let nh = Math.floor(total / 60) % 24;
    let nm = total % 60;

    const nap = nh >= 12 ? "PM" : "AM";
    nh = nh % 12 || 12;

    return `${String(nh).padStart(2, "0")}:${String(nm).padStart(2, "0")} ${nap}`;
}

function getGroupFromBatch(batch) {
    return batch.slice(0, -1); // "1A11" → "1A1"
}


function splitLastChar(str) {
    return {
        main: str.slice(0, -1),
        last: str.slice(-1)
    };
}

function getBatchFromIndex(i) {
    let current = null;

    const batches = Object.entries(batchesIndexed)
        .map(([name, obj]) => ({ name, indx: obj.indx }))
        .sort((a, b) => a.indx - b.indx);

    for (const b of batches) {
        if (i >= b.indx) {
            current = b.name;
        } else {
            break;
        }
    }

    return current;
}
