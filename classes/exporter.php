<?php
// This file is part of Moodle - http://moodle.org/
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
 * @copyright  2025 Your Name
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */

defined('MOODLE_INTERNAL') || die();

global $CFG;
require_once($CFG->libdir . '/gradelib.php');
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

/**
 * Custom Excel grade export class.
 *
 * @package    gradeexport_customexcel
 * @copyright  2025 Your Name
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */
class grade_export_customexcel extends grade_export {

    /**
     * Constructor.
     *
     * @param stdClass $course The course.
     * @param int $groupid The group ID.
     * @param string|null $itemlist Item list.
     * @param bool $exportfeedback Whether to export feedback.
     * @param bool $onlyactive Whether to include only active users.
     */

    /**
     * Generate and output the Excel file with grades.
     *
     * @return void
     */
    public function print_grades() {
        global $DB;

        $filename = clean_filename("grades-{$this->course->shortname}.xlsx");

        // Create PhpSpreadsheet workbook and worksheet.
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
        $weightstyle = ['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]];
        $notesstyle = ['font' => ['italic' => true]];
        $notesboldstyle = ['font' => ['bold' => true]];

        // Metadata and notes.
        $sheet->setCellValue('A1', 'Subject code');
        $sheet->setCellValue('B1', $this->course->shortname);

        $sheet->setCellValue('A2', 'Subject name');
        $sheet->setCellValue('B2', $this->course->fullname);

        $sheet->setCellValue('A3', 'Delivery mode');
        $sheet->setCellValue('B3', '---');

        $sheet->setCellValue('A4', 'Campus');
        $sheet->setCellValue('B4', '---');

        $sheet->setCellValue('D1', 'Please note:');
        $sheet->getStyle('D1')->applyFromArray($headerstyle);
        $sheet->setCellValue(
            'D2',
            'A dash (-) signifies a student that they did not submit the assessment and automatically fail the subject.'
        );
        $sheet->getStyle('D2')->applyFromArray($notesstyle);
        $sheet->setCellValue(
            'D3',
            'A zero (0) signifies a student has submitted an assessment but it was beyond the 2 week late assessment '
            . 'submission. They are still eligible to pass the subject if their overall total is greater than 50%.'
        );
        $sheet->getStyle('D3')->applyFromArray($notesstyle);
        $sheet->setCellValue('D4', 'All course totals are rounded to the whole number.');
        $sheet->getStyle('D4')->applyFromArray($notesboldstyle);
        // Fetch items and students.
        $items = grade_item::fetch_all(['courseid' => $this->course->id]);
        $context = context_course::instance($this->course->id);

        $studentrole = $DB->get_record('role', ['shortname' => 'student']);
        $users = [];
        if ($studentrole) {
            $users = get_role_users($studentrole->id, $context);
        }

        // Sort users by student number (idnumber).
        usort($users, function($a, $b) {
            return strcmp((string)$a->idnumber, (string)$b->idnumber);
        });
        // If no students → write message and exit.
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

        // Header row 1 (assessments, total, grade).
        $row = 6;  // Start at row 6.
        $col = 4; // Column D.
        $assessmentitems = [];
        foreach ($items as $item) {
            if ($item->itemtype === 'mod') {
                $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
                $sheet->setCellValue($coord, $item->get_name());
                $sheet->getStyle($coord)->applyFromArray($headerstyle);
                $assessmentitems[] = $item;
                $col++;
            }
        }
        $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
        $sheet->setCellValue($coord, 'Total');
        $sheet->getStyle($coord)->applyFromArray($headerstyle);
        $col++;
        $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
        $sheet->setCellValue($coord, 'Grade');
        $sheet->getStyle($coord)->applyFromArray($headerstyle);

        // Enable text wrapping on row 6 (assessment names row).
        $sheet->getStyle('A6:' . \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . '6')
            ->getAlignment()
            ->setWrapText(true)
            ->setVertical(Alignment::VERTICAL_CENTER);

        // Increase row height for visibility.
        $sheet->getRowDimension(6)->setRowHeight(15);
        // Header row 2 (Student ID + weightings).
        $row++;
        $sheet->setCellValue('A' . $row, 'Student ID');
        $sheet->getStyle('A' . $row)->applyFromArray($headerstyle);
        $sheet->setCellValue('B' . $row, 'First name');
        $sheet->getStyle('B' . $row)->applyFromArray($headerstyle);
        $sheet->setCellValue('C' . $row, 'Surname');
        $sheet->getStyle('C' . $row)->applyFromArray($headerstyle);

        $col = 4;
        $totalweight = 0;
        foreach ($assessmentitems as $item) {
            $weight = $item->aggregationcoef;
            $totalweight += $weight;
            $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
            $sheet->setCellValue($coord, $weight . '%');
            $sheet->getStyle($coord)->applyFromArray($weightstyle);
            $col++;
        }
        $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
        $sheet->setCellValue($coord, $totalweight . '%');
        $sheet->getStyle($coord)->applyFromArray($weightstyle);
        $col++;
        $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
        $sheet->setCellValue($coord, '');
        $sheet->getStyle($coord)->applyFromArray($weightstyle);
        // Set column widths.
        $sheet->getColumnDimension('A')->setWidth(14);
        $sheet->getColumnDimension('B')->setWidth(13);
        $sheet->getColumnDimension('C')->setWidth(13);
        for ($i = 4; $i <= $col; $i++) {
            $colletter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($i);
            $sheet->getColumnDimension($colletter)->setWidth(13);
        }
        // Student rows.
        $row++;
        $courseitem = grade_item::fetch(['courseid' => $this->course->id, 'itemtype' => 'course']);

        foreach ($users as $user) {
            $studentid = !empty($user->idnumber) ? $user->idnumber : '-';
            $sheet->setCellValue('A' . $row, $studentid);
            $sheet->setCellValue('B' . $row, $user->firstname);
            $sheet->setCellValue('C' . $row, $user->lastname);

            $col = 4;
            foreach ($assessmentitems as $item) {
                $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
                $grade = grade_grade::fetch(['itemid' => $item->id, 'userid' => $user->id]);
                $sheet->setCellValue($coord,
                    ($grade && $grade->finalgrade !== null) ? round($grade->finalgrade, 2) : '-'
                );
                $col++;
            }

            if ($courseitem) {
                $coursegrade = grade_grade::fetch(['itemid' => $courseitem->id, 'userid' => $user->id]);
                $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
                $sheet->setCellValue($coord,
                    ($coursegrade && $coursegrade->finalgrade !== null) ? round($coursegrade->finalgrade, 2) : '-'
                );
                $col++;
                $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
                $sheet->setCellValue($coord,
                    ($coursegrade && $coursegrade->finalgrade !== null)
                        ? grade_format_gradevalue_letter($coursegrade->finalgrade, $courseitem)
                        : '-'
                );
            }
            $row++;
        }

        // Output.

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header("Content-Disposition: attachment;filename=\"$filename\"");
        header('Cache-Control: max-age=0');

        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }
}
