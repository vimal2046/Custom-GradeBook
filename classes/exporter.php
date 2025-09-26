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
    $sheet->setCellValue('A2', 'Subject code');
    $sheet->setCellValue('B2', $this->course->shortname);
    $sheet->setCellValue('A3', 'Subject name');
    $sheet->setCellValue('B3', $this->course->fullname);
    // Auto-size metadata columns A and B
    $sheet->getColumnDimension('A')->setWidth(13);
    $sheet->getColumnDimension('B')->setWidth(13);
    $sheet->getStyle('A2')->getFont()->setBold(true); // Make label bold
    $sheet->getStyle('A3')->getFont()->setBold(true);

    $sheet->mergeCells('A5:b5');
    $sheet->setCellValue('A5', 'Reference Information:');
    $sheet->getStyle('A5')->applyFromArray($headerstyle);
    $sheet->setCellValue('B6', 'A dash (-) signifies no submission (automatic fail).');
    $sheet->getStyle('B6')->applyFromArray($notesstyle);
    $sheet->setCellValue('B7', 'A zero (0) signifies late submission beyond 2 weeks.');
    $sheet->getStyle('B7')->applyFromArray($notesstyle);
    $sheet->setCellValue('B8', 'The scores below are based on marks out of 100 for each assesment item. Weightings and grade scale are provided for reference');
    $sheet->getStyle('B8')->applyFromArray($notesstyle);
    $sheet->setCellValue('B9', 'All course totals are rounded to the whole number.');
    $sheet->getStyle('B9')->applyFromArray($notesboldstyle);


    // --------------------------------------------------------------------
    // Add Grade Letters reference (with fallback to system defaults).
    // --------------------------------------------------------------------
    $context = context_course::instance($this->course->id);
    $letters = grade_get_letters($context);

    // If no grade letters at course level → fallback to system defaults.
    if (empty($letters)) {
        $syscontext = context_system::instance();
        $letters = grade_get_letters($syscontext);
    }

    if (!empty($letters)) {
        // Start writing reference at column F row 2.
        $row = 2;
        $sheet->setCellValue("K$row", 'Highest');
        $sheet->setCellValue("L$row", 'Lowest');
        $sheet->setCellValue("M$row", 'Letter');
        $sheet->getStyle("K$row:M$row")->getFont()->setBold(true);

        // Loop through grade letters.
        $prevboundary = 100.00;
        foreach ($letters as $lowerboundary => $letter) {
            $row++;
            $sheet->setCellValue("K$row", $prevboundary . '%');
            $sheet->setCellValue("L$row", $lowerboundary . '%');
            $sheet->setCellValue("M$row", $letter);

            $prevboundary = $lowerboundary;
        }
    }


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
    }

    // 🚨 Final check: if no assignment items and no course total selected.
    if (empty($assessmentitems) && empty($courseitem)) {
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
$row = 18;
$col = 4;

// Student identity headers (A, B, C).
$sheet->setCellValue('A' . $row, 'Student ID');
$sheet->getStyle('A' . $row)->applyFromArray($headerstyle);

$sheet->setCellValue('B' . $row, 'First name');
$sheet->getStyle('B' . $row)->applyFromArray($headerstyle);

$sheet->setCellValue('C' . $row, 'Surname');
$sheet->getStyle('C' . $row)->applyFromArray($headerstyle);

foreach ($assessmentitems as $item) {
    $coord = Coordinate::stringFromColumnIndex($col) . $row;
    $sheet->setCellValue($coord, $item->get_name());
    $sheet->getStyle($coord)->applyFromArray($headerstyle);

    // ✅ Force width + wrap for this column
    $colletter = Coordinate::stringFromColumnIndex($col);
    $sheet->getColumnDimension($colletter)->setAutoSize(false);
    $sheet->getColumnDimension($colletter)->setWidth(15);
    $sheet->getStyle($colletter . $row)->getAlignment()
        ->setWrapText(true)
        ->setHorizontal(Alignment::HORIZONTAL_CENTER)
        ->setVertical(Alignment::VERTICAL_CENTER);

    $col += count($this->displaytype);

    if ($this->export_feedback) {
        $coord = Coordinate::stringFromColumnIndex($col) . $row;
        $sheet->setCellValue($coord, get_string('feedback'));
        $sheet->getStyle($coord)->applyFromArray($headerstyle);

        // ✅ Also apply for feedback column
        $colletter = Coordinate::stringFromColumnIndex($col);
        $sheet->getColumnDimension($colletter)->setAutoSize(false);
        $sheet->getColumnDimension($colletter)->setWidth(15);
        $sheet->getStyle($colletter . $row)->getAlignment()
            ->setWrapText(true)
            ->setHorizontal(Alignment::HORIZONTAL_CENTER)
            ->setVertical(Alignment::VERTICAL_CENTER);

        $col++;
    }
}

if ($courseitem) {
    foreach ($this->displaytype as $gradedisplayname => $gradedisplayconst) {
        $coord = Coordinate::stringFromColumnIndex($col) . $row;
        $sheet->setCellValue($coord, get_string($gradedisplayname, 'grades'));
        $sheet->getStyle($coord)->applyFromArray($headerstyle);

        // ✅ Force width + wrap for course total columns
        $colletter = Coordinate::stringFromColumnIndex($col);
        $sheet->getColumnDimension($colletter)->setAutoSize(false);
        $sheet->getColumnDimension($colletter)->setWidth(15);
        $sheet->getStyle($colletter . $row)->getAlignment()
            ->setWrapText(true)
            ->setHorizontal(Alignment::HORIZONTAL_CENTER)
            ->setVertical(Alignment::VERTICAL_CENTER);

        $col++;
    }
}

// Allow row 18 height to adjust automatically
$sheet->getRowDimension(18)->setRowHeight(-1);


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
