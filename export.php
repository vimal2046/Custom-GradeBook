<?php
defined('MOODLE_INTERNAL') || die();

require_once($CFG->libdir . '/gradelib.php');
require_once($CFG->libdir . '/excellib.class.php'); // Moodle Excel lib.
require_once($CFG->dirroot . '/grade/export/lib.php'); // Base grade_export class.

/**
 * Custom Excel grade export.
 */
class grade_export_customexcel extends grade_export {

    public function __construct($course, $groupid = 0, $itemlist = null, $exportfeedback = false, $onlyactive = false) {
        parent::__construct($course, $groupid, $itemlist, $exportfeedback, $onlyactive);
    }

    /**
     * Outputs the grade export in custom Excel format.
     */
    public function print_grades() {
        global $OUTPUT;

        $filename = clean_filename("grades-{$this->course->shortname}.xlsx");

        // Create workbook and worksheet.
        $workbook = new \MoodleExcelWorkbook('-');
        $workbook->send($filename);
        $worksheet = $workbook->add_worksheet('Grades');

        // -------------------------------
        // Write headerss
        // -------------------------------
        $headers = ['Student ID', 'Full Name'];
        foreach ($this->columns as $col) {
            $headers[] = $col->get_name();
        }

        $col = 0;
        foreach ($headers as $header) {
            $worksheet->write(0, $col, $header);
            $col++;
        }

        // -------------------------------
        // Write grade data
        // -------------------------------
        $row = 1;
        foreach ($this->grades as $userid => $usergrades) {
            $user = $this->users[$userid];

            $data = [
                $user->id,
                fullname($user)
            ];

            foreach ($this->columns as $col) {
                $data[] = $usergrades[$col->get_itemid()];
            }

            $col = 0;
            foreach ($data as $cell) {
                $worksheet->write($row, $col, $cell);
                $col++;
            }
            $row++;
        }

        // Finish export and trigger download.
        $workbook->close();
        exit;
    }
}
