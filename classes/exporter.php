<?php
// This file is part of Moodle - http://moodle.org/.
//
// Moodle is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// Moodle is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with Moodle.  If not, see <http://www.gnu.org/licenses/>.

/**
 * Custom Excel grade exporter.
 *
 * @package    gradeexport_customexcel
 */

defined('MOODLE_INTERNAL') || die();

require_once($CFG->dirroot . '/grade/export/lib.php');
require_once($CFG->dirroot . '/grade/lib.php');
// Load PhpSpreadsheet via composer vendor autoloader.
$vendorautoloader = __DIR__ . '/../vendor/autoload.php';
if (file_exists($vendorautoloader)) {
    require_once($vendorautoloader);
}

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
/**
 * Custom Excel grade export class.
 *
 * @package    gradeexport_customexcel
 * @copyright  2025 Your Name
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */
class grade_export_customexcel extends grade_export {
    public $plugin = 'customexcel';
    protected $formdata;

    public function __construct($course, $groupid, $formdata) {
        parent::__construct($course, $groupid, $formdata);
        $this->formdata = $formdata;
    }

    /**
     * Generate and output the Excel file with grades.
     */
    public function print_grades() {
    global $DB;

    $filename = clean_filename("grades-{$this->course->shortname}.xlsx");

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle('Results template sample');

    // Styles.
    $headerstyle = [
        'font' => ['bold' => true],
        'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER],
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'startColor' => ['rgb' => 'F8CBAD'],
        ],
    ];
    $notesstyle = ['font' => ['italic' => true]];
    $notesboldstyle = ['font' => ['bold' => true]];

    // Metadata and notes.
    $sheet->setCellValue('A1', 'Subject code');
    $sheet->setCellValue('B1', $this->course->shortname);
    $sheet->setCellValue('A2', 'Subject name');
    $sheet->setCellValue('B2', $this->course->fullname);

    $sheet->setCellValue('D1', 'Please note:');
    $sheet->getStyle('D1')->applyFromArray($headerstyle);
    $sheet->setCellValue('D2', 'A dash (-) signifies no submission (automatic fail).');
    $sheet->getStyle('D2')->applyFromArray($notesstyle);
    $sheet->setCellValue('D3', 'A zero (0) signifies late submission beyond 2 weeks.');
    $sheet->getStyle('D3')->applyFromArray($notesstyle);
    $sheet->setCellValue('D4', 'All course totals are rounded to the whole number.');
    $sheet->getStyle('D4')->applyFromArray($notesboldstyle);

    // --------------------------------------------------------------------
    // Handle selected grade items.
    // --------------------------------------------------------------------
    $selecteditemids = [];
    if (!empty($this->formdata->itemids)) {
        if (is_array($this->formdata->itemids)) {
            $selecteditemids = $this->formdata->itemids;
        } else {
            $selecteditemids = explode(',', $this->formdata->itemids);
        }
    }

    $assessmentitems = [];
    $courseitem = null;

    if (!empty($selecteditemids)) {
        foreach ($selecteditemids as $itemid) {
            $item = grade_item::fetch(['id' => $itemid, 'courseid' => $this->course->id]);
            if ($item && $item->itemtype === 'mod') {
                $assessmentitems[] = $item;
            }
            if ($item && $item->itemtype === 'course') {
                $courseitem = $item;
            }
        }
    } else {
        // 🚨 No items selected → stop and output message.
        $sheet->setCellValue('A6', 'No grade items selected for export.');
        $sheet->getStyle('A6')->applyFromArray([
            'font' => ['bold' => true, 'italic' => true, 'color' => ['rgb' => 'FF0000']],
        ]);
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header("Content-Disposition: attachment;filename=\"$filename\"");
        header('Cache-Control: max-age=0');
        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }

    // --------------------------------------------------------------------
    // Users (active only if selected).
    // --------------------------------------------------------------------
    $users = [];
    $gui = new graded_users_iterator($this->course, $assessmentitems, $this->groupid);
    $gui->require_active_enrolment($this->onlyactive);
    $gui->init();

    while ($userdata = $gui->next_user()) {
        $users[] = $userdata->user;
    }

    $gui->close();

    usort($users, fn($a, $b) => strcmp((string)$a->idnumber, (string)$b->idnumber));

    if (empty($users)) {
        $sheet->setCellValue('A6', 'No students are enrolled in this course or no grades available.');
        $sheet->getStyle('A6')->applyFromArray([
            'font' => ['bold' => true, 'italic' => true, 'color' => ['rgb' => 'FF0000']],
        ]);
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header("Content-Disposition: attachment;filename=\"$filename\"");
        header('Cache-Control: max-age=0');
        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }

    // --------------------------------------------------------------------
    // Header row 1 (assessment names).
    // --------------------------------------------------------------------
    $row = 6;
    $col = 4;
    foreach ($assessmentitems as $item) {
        $coord = Coordinate::stringFromColumnIndex($col) . $row;
        $sheet->setCellValue($coord, $item->get_name());
        $sheet->getStyle($coord)->applyFromArray($headerstyle);
        $col += count($this->displaytype);
        if ($this->export_feedback) {
            $coord = Coordinate::stringFromColumnIndex($col) . $row;
            $sheet->setCellValue($coord, get_string('feedback'));
            $sheet->getStyle($coord)->applyFromArray($headerstyle);
            $col++;
        }
    }

    if ($courseitem) {
        foreach ($this->displaytype as $gradedisplayname => $gradedisplayconst) {
            $coord = Coordinate::stringFromColumnIndex($col) . $row;
            $sheet->setCellValue($coord, get_string($gradedisplayname, 'grades'));
            $col++;
        }
    }

    // --------------------------------------------------------------------
    // Student rows.
    // --------------------------------------------------------------------
    $row++;
    foreach ($users as $user) {
        $c = 1;
        $sheet->setCellValue(Coordinate::stringFromColumnIndex($c++) . $row, $user->idnumber ?: '-');
        $sheet->setCellValue(Coordinate::stringFromColumnIndex($c++) . $row, $user->firstname);
        $sheet->setCellValue(Coordinate::stringFromColumnIndex($c++) . $row, $user->lastname);

        foreach ($assessmentitems as $item) {
            $grade = grade_grade::fetch(['itemid' => $item->id, 'userid' => $user->id]);
            foreach ($this->displaytype as $gradedisplayconst) {
                $val = ($grade && $grade->finalgrade !== null)
                    ? $this->format_grade($grade, $gradedisplayconst)
                    : '-';
                $sheet->setCellValue(Coordinate::stringFromColumnIndex($c++) . $row, $val);
            }
            if ($this->export_feedback) {
                $feedbacktext = ($grade && !empty(trim(strip_tags($grade->feedback))))
                    ? trim(strip_tags($grade->feedback))
                    : '-';
                $sheet->setCellValue(Coordinate::stringFromColumnIndex($c++) . $row, $feedbacktext);
            }
        }

        if ($courseitem) {
            $coursegrade = grade_grade::fetch(['itemid' => $courseitem->id, 'userid' => $user->id]);
            foreach ($this->displaytype as $gradedisplayconst) {
                $val = ($coursegrade && $coursegrade->finalgrade !== null)
                    ? $this->format_grade($coursegrade, $gradedisplayconst)
                    : '-';
                $sheet->setCellValue(Coordinate::stringFromColumnIndex($c++) . $row, $val);
            }
        }
        $row++;
    }

    // --------------------------------------------------------------------
    // Output file.
    // --------------------------------------------------------------------
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header("Content-Disposition: attachment;filename=\"$filename\"");
    $writer = new Xlsx($spreadsheet);
    $writer->save('php://output');
    exit;
}


}
