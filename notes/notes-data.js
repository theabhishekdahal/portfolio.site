/**
 * Notes Registry
 * Add an entry here for every HTML note you upload to the /notes/ folder.
 *
 * Fields:
 *   id          – unique slug, used as the `?note=<id>` URL parameter
 *   title       – display title
 *   description – short summary shown on the listing card
 *   date        – ISO date string (YYYY-MM-DD)
 *   category    – used for the colour-coded badge
 *   tags        – array of keyword strings
 *   file        – filename of the HTML note (relative to /notes/)
 */
const notesData = [
    {
        id: "cap1-financial-accounting",
        title: "CAP I – Financial Accounting: Key Concepts",
        description: "Core double-entry principles, trial balance, adjusting entries, and financial statement preparation covered in CAP I.",
        date: "2025-03-01",
        category: "Financial Accounting",
        tags: ["CAP I", "Double Entry", "Trial Balance", "Financial Statements"],
        file: "cap1-financial-accounting.html"
    }
    {
        id: "cap1-financial-accounting",
        title: "Mutual Funds",
       date: "2025-03-01",
        category: "Finance",
        file: "Mutual Fund.html"
    }
];
