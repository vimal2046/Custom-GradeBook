<?php
defined('MOODLE_INTERNAL') || die();

require_once($CFG->libdir . '/gradelib.php');
require_once($CFG->libdir . '/excellib.class.php');
require_once($CFG->dirroot . '/grade/export/lib.php');

/**
 * Custom Excel grade export.
 */
class grade_export_customexcel extends grade_export {

    public function __construct($course, $groupid = 0, $itemlist = null, $exportfeedback = false, $onlyactive = false) {
        parent::__construct($course, $groupid, $itemlist, $exportfeedback, $onlyactive);
    }

    public function print_grades() {
        global $CFG;

        $filename = clean_filename("grades-{$this->course->shortname}.xlsx");

        // Workbook setup.
        $workbook = new \MoodleExcelWorkbook('-');
        $workbook->send($filename);
        $worksheet = $workbook->add_worksheet('Results template sample');

        // Formats.
        $format_bold = $workbook->add_format(['bold' => 1]);
        $format_notes = $workbook->add_format(['italic' => 1]);

        // Metadata.
        $worksheet->write(0, 0, 'Subject code');
        $worksheet->write(0, 1, $this->course->shortname);
        $worksheet->write(1, 0, 'Subject name');
        $worksheet->write(1, 1, $this->course->fullname);
        $worksheet->write(2, 0, 'Delivery mode');
        $worksheet->write(2, 1, '---');
        $worksheet->write(3, 0, 'Campus');
        $worksheet->write(3, 1, '---');

        // Notes.
        $worksheet->write(0, 3, 'Please note:', $format_bold);
        $worksheet->write(1, 3, 'A dash (-) signifies a student did not attempt the task', $format_notes);
        $worksheet->write(2, 3, 'A zero (0) signifies a student submitted but got 0 marks', $format_notes);
        $worksheet->write(3, 3, 'All Course totals are rounded to the whole number', $format_notes);

        // Header row.
        $row = 6;
        $col = 0;
        $worksheet->write($row, $col++, 'Student ID', $format_bold);
        $worksheet->write($row, $col++, 'First name', $format_bold);
        $worksheet->write($row, $col++, 'Surname', $format_bold);

        foreach ($this->columns as $colitem) {
            $worksheet->write($row, $col++, $colitem->get_name(), $format_bold);
        }

        $worksheet->write($row, $col++, 'Total', $format_bold);
        $worksheet->write($row, $col++, 'Grade', $format_bold);

        // Weighting row.
        $row++;
        $col = 3;
        foreach ($this->columns as $colitem) {
            $worksheet->write($row, $col, $colitem->get_item()->aggregationcoef, $format_notes);
            $col++;
        }

        // Student rows.
        $row++;
        foreach ($this->users as $userid => $user) {
            $col = 0;

            $studentid = $user->idnumber ?: $user->id;
            $worksheet->write($row, $col++, $studentid);
            $worksheet->write($row, $col++, $user->firstname);
            $worksheet->write($row, $col++, $user->lastname);

            foreach ($this->columns as $colitem) {
                $val = $this->grades[$userid][$colitem->get_itemid()] ?? null;
                $worksheet->write($row, $col++, ($val !== null ? round($val, 2) : '-'));
            }

            // Course total and letter grade.
            $final = grade_get_course_grades($this->course->id, null, $userid);
            if (!empty($final->grades[$userid])) {
                $worksheet->write($row, $col++, round($final->grades[$userid]->grade, 0));
                $worksheet->write($row, $col++, $final->grades[$userid]->str_grade);
            } else {
                $worksheet->write($row, $col++, '-');
                $worksheet->write($row, $col++, '-');
            }

            $row++;
        }

        $workbook->close();
        exit;
    }
}
