//! #1969: a wrapped cell's row must fit WORD-wrapped lines, not the character-fit
//! estimate (words can't split, so each drawn line ends early and narrow columns
//! need more lines than width/estimate suggests).

use gridcraft_engine::Session;
use serde_json::json;

fn session() -> Session {
    let mut s = Session::new();
    s.new_workbook();
    s
}

fn row_height(s: &mut Session, row: u32) -> f32 {
    s.doc().unwrap().wb.active().unwrap().row_height(row)
}

#[test]
fn wrapped_row_fits_word_wrapped_lines() {
    let mut s = session();
    // Column fits ~10 characters; six-character words mean only one word per line.
    s.execute("home.columnWidth", json!({"cols": "A:A", "chars": 10})).unwrap();
    s.execute("cell.set", json!({"cell": "A1", "input": "aaaaaa bbbbbb cccccc dddddd"})).unwrap();
    let plain = row_height(&mut s, 0);
    s.execute("home.wrapText", json!({"range": "A1", "on": true})).unwrap();
    let wrapped = row_height(&mut s, 0);
    // Word wrap: 4 lines (one word each). The old estimate: ceil(28/10) = 3 lines.
    // Line height at 11pt is ~19.8, so 4 lines need ~79 vs ~59 for 3.
    assert!(wrapped >= plain + 3.0 * 19.0, "4 word-wrapped lines must grow the row by ~4 line heights: plain={plain}, wrapped={wrapped}");
}

#[test]
fn issue_repro_nasa_sentence() {
    let mut s = session();
    s.execute("home.columnWidth", json!({"cols": "A:A", "chars": 12})).unwrap();
    s.execute("cell.set", json!({"cell": "A1", "input": "Since the 1960s, NASA has left hardware across the solar system\u{2013}on the Moon, on Mars, and in deep space."}))
        .unwrap();
    s.execute("home.wrapText", json!({"range": "A1", "on": true})).unwrap();
    let h = row_height(&mut s, 0);
    // Excel shows this at ~10 lines in a 12-char column; the character estimate undercounts.
    assert!(h >= 9.0 * 19.0, "NASA sentence in a 12-char column needs ~10 lines, got height {h}");
}

#[test]
fn khmer_rows_fit_their_stacked_lines() {
    // Khmer draws through a face whose line is about 2 em: a 15pt row cuts off the subscripts and
    // vowels that stack below the baseline, and wrapped lines would run into each other. Row
    // heights here are in pixels, like the font's em: 11pt is 14.67px, so a Khmer line is ~29px.
    let line = 2.0 * 11.0 * 96.0 / 72.0;
    let mut s = session();
    let default = row_height(&mut s, 0);
    s.execute("cell.set", json!({"cell": "A1", "input": "កម្ពុជា"})).unwrap();
    s.execute("cell.set", json!({"cell": "A2", "input": "Cambodia"})).unwrap();
    let khmer = row_height(&mut s, 0);
    assert!(khmer >= line, "an 11pt Khmer line needs about {line}px, got {khmer}");
    assert_eq!(row_height(&mut s, 1), default, "Latin text keeps the default row");

    s.execute("home.columnWidth", json!({"cols": "B:B", "chars": 10})).unwrap();
    s.execute("cell.set", json!({"cell": "B3", "input": "ព្រះរាជាណាចក្រកម្ពុជា ជាតិ សាសនា ព្រះមហាក្សត្រ"})).unwrap();
    s.execute("home.wrapText", json!({"range": "B3", "on": true})).unwrap();
    let wrapped = row_height(&mut s, 2);
    // The 21-character first phrase has no spaces and breaks between clusters over three lines.
    assert!(wrapped >= 6.0 * line, "six or more wrapped Khmer lines of about {line}px each, got {wrapped}");

    // Replacing the Khmer gives the row back its default height.
    s.execute("cell.set", json!({"cell": "A1", "input": "Phnom Penh"})).unwrap();
    assert_eq!(row_height(&mut s, 0), default);
}
