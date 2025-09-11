<?php
defined('MOODLE_INTERNAL') || die();

require_once($CFG->libdir . '/gradelib.php');
require_once($CFG->libdir . '/excellib.class.php');
require_once($CFG->dirroot . '/grade/export/lib.php');
require_once($CFG->dirroot . '/grade/lib.php');

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

        // -------------------------------
        // Metadata & Notes
        // -------------------------------
        $worksheet->write(0, 0, 'Subject code');
        $worksheet->write(0, 1, $this->course->shortname);

        $worksheet->write(1, 0, 'Subject name');
        $worksheet->write(1, 1, $this->course->fullname);

        $worksheet->write(2, 0, 'Delivery mode');
        $worksheet->write(2, 1, '---'); // TODO: fetch from course custom fields if available.

        $worksheet->write(3, 0, 'Campus');
        $worksheet->write(3, 1, '---');

        // Notes (right side).
        $worksheet->write(0, 3, 'Please note:', $format_bold);
        $worksheet->write(1, 3, 'A dash (-) signifies a student did not attempt the task', $format_notes);
        $worksheet->write(2, 3, 'A zero (0) signifies a student submitted but got 0 marks', $format_notes);
        $worksheet->write(3, 3, 'All Course totals are rounded to the whole number', $format_notes);

        // -------------------------------
        // Fetch assessment items and students
        // -------------------------------
        $items = grade_item::fetch_all(['courseid' => $this->course->id]);
        $context = context_course::instance($this->course->id);
        $users = get_enrolled_users($context);

        // -------------------------------
        // Header row
        // -------------------------------
        $row = 6;
        $col = 0;
        $worksheet->write($row, $col++, 'Student ID', $format_bold);
        $worksheet->write($row, $col++, 'First name', $format_bold);
        $worksheet->write($row, $col++, 'Surname', $format_bold);

        $assessmentitems = [];
        foreach ($items as $item) {
            if ($item->itemtype === 'mod') {
                $worksheet->write($row, $col++, $item->get_name(), $format_bold);
                $assessmentitems[] = $item;
            }
        }

        $worksheet->write($row, $col++, 'Total', $format_bold);
        $worksheet->write($row, $col++, 'Grade', $format_bold);

        // -------------------------------
        // Weighting row
        // -------------------------------
        $row++;
        $col = 3;
        foreach ($assessmentitems as $item) {
            $worksheet->write($row, $col++, $item->aggregationcoef, $format_notes);
        }

        // -------------------------------
        // Student rows
        // -------------------------------
        $row++;
        $courseitem = grade_item::fetch(['courseid' => $this->course->id, 'itemtype' => 'course']);

        foreach ($users as $user) {
            $col = 0;

            $studentid = $user->idnumber ?: $user->id;
            $worksheet->write($row, $col++, $studentid);
            $worksheet->write($row, $col++, $user->firstname);
            $worksheet->write($row, $col++, $user->lastname);

            // Grades per assessment
            foreach ($assessmentitems as $item) {
                $grade = grade_grade::fetch(['itemid' => $item->id, 'userid' => $user->id]);
                if ($grade && $grade->finalgrade !== null) {
                    $worksheet->write($row, $col++, round($grade->finalgrade, 2));
                } else {
                    $worksheet->write($row, $col++, '-');
                }
            }

            // Course total + Letter grade
            if ($courseitem) {
                $coursegrade = grade_grade::fetch(['itemid' => $courseitem->id, 'userid' => $user->id]);
                if ($coursegrade && $coursegrade->finalgrade !== null) {
                    $worksheet->write($row, $col++, round($coursegrade->finalgrade, 0));
                    $worksheet->write($row, $col++, grade_format_gradevalue_letter($coursegrade->finalgrade, $courseitem));
                } else {
                    $worksheet->write($row, $col++, '-');
                    $worksheet->write($row, $col++, '-');
                }
            } else {
                $worksheet->write($row, $col++, '-');
                $worksheet->write($row, $col++, '-');
            }

            $row++;
        }

        // -------------------------------
        // Close file
        // -------------------------------
        $workbook->close();
        exit;
    }
}
