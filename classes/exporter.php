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
 * @copyright  2025 AC University
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
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
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

/**
 * Custom Excel grade export class.
 *
 * @package    gradeexport_customexcel
 * @copyright  2025 Your Name
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */
class grade_export_customexcel extends grade_export {
    /**
     * @var string Plugin name identifier.
     */
    public $plugin = 'customexcel';
    /**
     * @var stdClass Form data submitted from export form.
     */
    protected $formdata;
    /**
     * Constructor for custom Excel grade export.
     *
     * @param stdClass $course  Course object.
     * @param int $groupid      Group ID for export filtering.
     * @param stdClass $formdata Form submission data from export form.
     */
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

        // Adding logo.

        // Insert logo in first row (A1).
        $logo = new Drawing();
        $logo->setName('Logo');
        $logo->setDescription('Institution Logo');

        // Path: plugin root (since exporter.php is in classes/, go up one folder).
        $logo->setPath(__DIR__ . '/../logo.png');  // Adjust filename if different.
        $logo->setHeight(60); // Adjust logo height.
        $logo->setCoordinates('A1'); // Place at cell A1.
        $logo->setOffsetX(5);  // Small horizontal offset.
        $logo->setOffsetY(5);  // Small vertical offset.
        $logo->setWorksheet($sheet);

        $sheet->getRowDimension(1)->setRowHeight(50);
        // Styles.
        $headerstyle = [
            'font' => ['bold' => true],
            'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => 'F8CBAD'],
            ],
        ];

        $studentinfostyle = [
            'font' => ['bold' => true, 'color' => ['rgb' => '000000']],
            'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => '538AC8'], // Blue.
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ],
            ],
        ];

        $assessmentstyle = [
            'font' => ['bold' => true],
            'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => 'FFF1E3'], // Light orange.
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ],
            ],
        ];

        $coursetotalstyle = [
            'font' => ['bold' => true, 'color' => ['rgb' => '00000']],
            'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => '538AC8'], // Blue.
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ],
            ],
        ];

        $refstyle = [
            'font' => ['bold' => true, 'color' => ['rgb' => 'FFFFFF']],
            'alignment' => ['horizontal' => Alignment::HORIZONTAL_LEFT, 'vertical' => Alignment::VERTICAL_CENTER],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => '131346'], // Dark navy.
            ],
        ];

        $gradescalestyle = [
            'font' => ['bold' => true, 'color' => ['rgb' => 'FFFFFF']],
            'alignment' => ['horizontal' => Alignment::HORIZONTAL_LEFT, 'vertical' => Alignment::VERTICAL_CENTER],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => '131346'], // Dark navy.
            ],
        ];

        $notesstyle = ['font' => ['italic' => true]];
        $notesboldstyle = ['font' => ['bold' => true]];

        // Metadata and notes.
        $sheet->setCellValue('A2', 'Subject code');
        $sheet->setCellValue('B2', $this->course->shortname);
        $sheet->setCellValue('A3', 'Subject name');
        $sheet->setCellValue('B3', $this->course->fullname);
        // Auto-size metadata columns A and B.
        $sheet->getColumnDimension('A')->setWidth(13);
        $sheet->getColumnDimension('B')->setWidth(13);
        $sheet->getStyle('A2')->getFont()->setBold(true);
        $sheet->getStyle('A3')->getFont()->setBold(true);

        $sheet->mergeCells('A5:B5');
        $sheet->setCellValue('A5', 'Reference Information:');
        $sheet->getStyle('A5')->applyFromArray($refstyle);
        $sheet->setCellValue('B6', 'A dash (-) signifies no submission (automatic fail).');
        $sheet->getStyle('B6')->applyFromArray($notesstyle);
        $sheet->setCellValue('B7', 'A zero (0) signifies late submission beyond 2 weeks.');
        $sheet->getStyle('B7')->applyFromArray($notesstyle);
        $sheet->setCellValue('B8',
            'The scores below are based on marks out of 100 for each assesment item.' .
            'Weightings and grade scale are provided for reference');
        $sheet->getStyle('B8')->applyFromArray($notesstyle);
        $sheet->setCellValue('B9', 'All course totals are rounded to the whole number.');
        $sheet->getStyle('B9')->applyFromArray($notesboldstyle);

        // Grade Letters reference (single-column scale in A11 / B11↓).
        $context = context_course::instance($this->course->id);
        $letters = grade_get_letters($context);
        if (empty($letters)) {
            $syscontext = context_system::instance();
            $letters = grade_get_letters($syscontext);
        }

        if (!empty($letters)) {
            // Place "Grade scale" at A11.
            $row = 11;
            $sheet->setCellValue("A{$row}", 'Grade scale:');
            $sheet->getStyle("A{$row}")->applyFromArray($gradescalestyle);;

            // Keep previous upper boundary (start at 100).
            $prevboundary = 100.0;

            // Small helper to nicely format numbers (no decimals when whole).
            $formatbound = function($v) {
                if (is_numeric($v) && floor($v) == $v) {
                    return (int)$v;
                }
                return number_format($v, 2);
            };

            foreach ($letters as $lowerboundary => $letter) {
                // Grade strings go into column B (starting same row).
                $sheet->setCellValue("B{$row}", "{$letter}: " .
                    $formatbound((float)$lowerboundary) . '-' . $formatbound($prevboundary));
                // Next row for next grade.
                $row++;
                $prevboundary = (float)$lowerboundary;
            }
        }

        // Handle selected grade items.
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

                if ($item) {
                    if ($item->itemtype === 'mod') {
                        // Regular assignment.
                        $assessmentitems[] = $item;
                    } else if ($item->itemtype === 'category') {
                        // Category total (e.g. Assessment 2 group).
                        $assessmentitems[] = $item;
                    } else if ($item->itemtype === 'course') {
                        // Course total.
                        $courseitem = $item;
                    }
                }
            }
        }

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

        // Users.
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

        // Header row.
        $row = 18;
        $col = 4;

        // Border style for headers.
        $borderstyle = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ],
            ],
        ];

        // Student identity headers.
        $sheet->setCellValue('A' . $row, 'Student ID');
        $sheet->getStyle('A' . $row)->applyFromArray($studentinfostyle);
        $sheet->getColumnDimension('A')->setWidth(15);
        $sheet->getStyle('A' . $row)->getAlignment()->setWrapText(true);

        $sheet->setCellValue('B' . $row, 'First name');
        $sheet->getStyle('B' . $row)->applyFromArray($studentinfostyle);
        $sheet->getColumnDimension('B')->setWidth(15);
        $sheet->getStyle('B' . $row)->getAlignment()->setWrapText(true);
        $sheet->setCellValue('C' . $row, 'Surname');
        $sheet->getStyle('C' . $row)->applyFromArray($studentinfostyle);
        $sheet->getColumnDimension('C')->setWidth(15);
        $sheet->getStyle('C' . $row)->getAlignment()->setWrapText(true);

        // Assignments.
        foreach ($assessmentitems as $item) {
            $startcol = $col;

            // Subcolumns (Real, Percentage, Letter, etc.).
            foreach ($this->displaytype as $gradedisplayname => $gradedisplayconst) {
                $coord = Coordinate::stringFromColumnIndex($col) . $row;
                $sheet->setCellValue($coord, get_string($gradedisplayname, 'grades'));
                $sheet->getStyle($coord)->applyFromArray($assessmentstyle);

                // Apply fixed width + wrap.
                $letter = Coordinate::stringFromColumnIndex($col);
                $sheet->getColumnDimension($letter)->setWidth(15);
                $sheet->getStyle($coord)->getAlignment()->setWrapText(true);

                $col++;
            }

            // Merge across subcolumns → assignment name.
            $startcolletter = Coordinate::stringFromColumnIndex($startcol);
            $endcolletter   = Coordinate::stringFromColumnIndex($col - 1);
            $sheet->mergeCells("{$startcolletter}{$row}:{$endcolletter}{$row}");
             // Shows Assignment name + total as a category total name.
            $displayname = $item->get_name();

            if ($item->itemtype === 'category') {
                $cat = grade_category::fetch(['id' => $item->iteminstance]);
                if ($cat) {
                    $displayname = $cat->fullname . ' total'; // E.g. "Assessment 1 total".
                }
            }

            $sheet->setCellValue($startcolletter . $row, $displayname);

            $sheet->getStyle("{$startcolletter}{$row}:{$endcolletter}{$row}")
                ->applyFromArray($assessmentstyle)
                ->applyFromArray($borderstyle)
                ->getAlignment()->setWrapText(true);

            if ($this->export_feedback) {
                $coord = Coordinate::stringFromColumnIndex($col) . $row;
                $sheet->setCellValue($coord, get_string('feedback'));
                $sheet->getStyle($coord)->applyFromArray($headerstyle)->applyFromArray($borderstyle);
                $letter = Coordinate::stringFromColumnIndex($col);
                $sheet->getColumnDimension($letter)->setWidth(15);
                $sheet->getStyle($coord)->getAlignment()->setWrapText(true);
                $col++;
            }
        }

        if ($courseitem) {
            $startcol = $col;
            foreach ($this->displaytype as $gradedisplayname => $gradedisplayconst) {
                $coord = Coordinate::stringFromColumnIndex($col) . $row;
                $sheet->setCellValue($coord, get_string($gradedisplayname, 'grades'));
                $sheet->getStyle($coord)->applyFromArray($headerstyle)->applyFromArray($borderstyle);

                // Apply fixed width + wrap.
                $letter = Coordinate::stringFromColumnIndex($col);
                $sheet->getColumnDimension($letter)->setWidth(15);
                $sheet->getStyle($coord)->getAlignment()->setWrapText(true);

                $col++;
            }
            $startcolletter = Coordinate::stringFromColumnIndex($startcol);
            $endcolletter   = Coordinate::stringFromColumnIndex($col - 1);
            $sheet->mergeCells("{$startcolletter}{$row}:{$endcolletter}{$row}");
            $sheet->setCellValue($startcolletter . $row, get_string('coursetotal', 'grades'));
            $sheet->getStyle("{$startcolletter}{$row}:{$endcolletter}{$row}")
                ->applyFromArray($coursetotalstyle)
                ->getAlignment()->setWrapText(true);
        }

        // Grade column.
        $coord = Coordinate::stringFromColumnIndex($col) . $row;
        $sheet->setCellValue($coord, 'Grade');
        $sheet->getStyle($coord)->applyFromArray($studentinfostyle);
        $sheet->getColumnDimension(Coordinate::stringFromColumnIndex($col))->setWidth(18);
        $sheet->getStyle($coord)->getAlignment()->setWrapText(true);

        $gradecolindex = $col;
        $col++;

        $sheet->getRowDimension(18)->setRowHeight(-1);

        // Freeze the top 18 rows so headers stay visible when scrolling.
        $sheet->freezePane('A19');

        // Fix: Force wrap and auto-adjust row height manually.
        $highestcolumn = $sheet->getHighestColumn(18);
        $range = "A18:{$highestcolumn}18";

        // Ensure wrapping is turned on for all header cells.
        $sheet->getStyle($range)->getAlignment()->setWrapText(true);

        // Manually estimate row height based on longest text in the header row.
        $maxtextlength = 0;
        foreach (range('A', $highestcolumn) as $colletter) {
            $cellvalue = $sheet->getCell("{$colletter}18")->getValue();
            if (is_string($cellvalue)) {
                $length = strlen($cellvalue);
                if ($length > $maxtextlength) {
                    $maxtextlength = $length;
                }
            }
        }

        // Adjust row height based on text length (rough approximation).
        // You can tune the multiplier (0.8 or 1.2) if you want tighter spacing.
        $baseheight = 20; // Default single-line height.
        $extralines = ceil($maxtextlength / 25); // 25 chars per line at width=15 roughly.
        $rowheight = $baseheight + ($extralines * 12);

        $sheet->getRowDimension(18)->setRowHeight($rowheight + 4);


        // Student rows.
        $row++;
        foreach ($users as $user) {
            $c = 1;
            $nonsubmission = false; // Track missing submission per student.

            // Student ID.
            $coord = Coordinate::stringFromColumnIndex($c) . $row;
            $sheet->setCellValue($coord, $user->idnumber ?: '-');
            $sheet->getStyle($coord)->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                ->setVertical(Alignment::VERTICAL_CENTER);
            $c++;

            // First name.
            $coord = Coordinate::stringFromColumnIndex($c) . $row;
            $sheet->setCellValue($coord, $user->firstname);
            $sheet->getStyle($coord)->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                ->setVertical(Alignment::VERTICAL_CENTER);
            $c++;

            // Surname.
            $coord = Coordinate::stringFromColumnIndex($c) . $row;
            $sheet->setCellValue($coord, $user->lastname);
            $sheet->getStyle($coord)->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                ->setVertical(Alignment::VERTICAL_CENTER);
            $c++;

            // Assignments & category items.
            foreach ($assessmentitems as $item) {
                $grade = grade_grade::fetch(['itemid' => $item->id, 'userid' => $user->id]);
                foreach ($this->displaytype as $gradedisplayconst) {
                    $val = ($grade && $grade->finalgrade !== null)
                        ? $this->format_grade($grade, $gradedisplayconst)
                        : '-';

                    $coord = Coordinate::stringFromColumnIndex($c) . $row;
                    $sheet->setCellValue($coord, $val);

                    // Center align.
                    $sheet->getStyle($coord)->getAlignment()
                        ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                        ->setVertical(Alignment::VERTICAL_CENTER);

                    // If value is "-" → mark non-submission + red bg.
                    if ($val === '-') {
                        $nonsubmission = true;
                        $sheet->getStyle($coord)->getFill()->setFillType(Fill::FILL_SOLID)
                            ->getStartColor()->setRGB('EF4C4D'); // Light red background.
                    }

                    $c++;
                }
                if ($this->export_feedback) {
                    $feedbacktext = ($grade && !empty(trim(strip_tags($grade->feedback))))
                        ? trim(strip_tags($grade->feedback))
                        : '-';
                    $coord = Coordinate::stringFromColumnIndex($c) . $row;
                    $sheet->setCellValue($coord, $feedbacktext);
                    $c++;
                }
            }

            $finalpercent = null;
            if ($courseitem) {
                $coursegrade = grade_grade::fetch(['itemid' => $courseitem->id, 'userid' => $user->id]);
                foreach ($this->displaytype as $gradedisplayconst) {
                    $val = ($coursegrade && $coursegrade->finalgrade !== null)
                        ? $this->format_grade($coursegrade, $gradedisplayconst)
                        : '-';

                    $coord = Coordinate::stringFromColumnIndex($c) . $row;
                    $sheet->setCellValue($coord, $val);

                    // Center align.
                    $sheet->getStyle($coord)->getAlignment()
                        ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                        ->setVertical(Alignment::VERTICAL_CENTER);

                    // If missing → mark non-submission + red bg.
                    if ($val === '-') {
                        $nonsubmission = true;
                        $sheet->getStyle($coord)->getFill()->setFillType(Fill::FILL_SOLID)
                            ->getStartColor()->setRGB('EF4C4D');
                    }

                    $c++;
                }

                if ($coursegrade && $coursegrade->finalgrade !== null) {
                    $finalpercent = floatval($coursegrade->finalgrade / $courseitem->grademax * 100);
                }
            }

            // Final Grade.
            $gradecoord = Coordinate::stringFromColumnIndex($gradecolindex) . $row;
            if ($nonsubmission) {
                $sheet->setCellValue($gradecoord, 'Fail (Non submission)');
                $sheet->getStyle($gradecoord)->getFont()->getColor()->setRGB('FF0000'); // Red text.
                $sheet->getStyle($gradecoord)->getFill()->setFillType(Fill::FILL_SOLID)
                    ->getStartColor()->setRGB('FFCCCC'); // Light red background.
            } else {
                $sheet->setCellValue($gradecoord, $finalpercent !== null
                    ? $this->get_grade_letter($finalpercent)
                    : '-');
                $sheet->getStyle($gradecoord)->getAlignment()
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                    ->setVertical(Alignment::VERTICAL_CENTER);
            }

            $row++;
        }

        // Apply thin border to the entire used range.
        $lastcol = Coordinate::stringFromColumnIndex($col - 1);
        $lastrow = $row - 1; // Because loop already incremented after last student.

        $sheet->getStyle("A18:{$lastcol}{$lastrow}")->applyFromArray([
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ],
            ],
        ]);

        // Output.
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header("Content-Disposition: attachment;filename=\"$filename\"");
        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }

    /**
     * Converts a numeric percentage into a letter grade based on course grade letters.
     *
     * @param float $percentage The numeric grade percentage.
     * @return string The corresponding grade letter or '-' if not found.
     */
    protected function get_grade_letter($percentage) {
        $context = context_course::instance($this->course->id);
        $letters = grade_get_letters($context);
        if (empty($letters)) {
            $letters = grade_get_letters(context_system::instance());
        }
        foreach ($letters as $lowerboundary => $letter) {
            if ($percentage >= $lowerboundary) {
                return $letter;
            }
        }
        return '-';
    }
}


