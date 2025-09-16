<?php
defined('MOODLE_INTERNAL') || die();

require_once($CFG->libdir . '/gradelib.php');
require_once($CFG->dirroot . '/grade/export/lib.php');
require_once($CFG->dirroot . '/grade/lib.php');
require_once(__DIR__ . '/vendor/autoload.php');

// PhpSpreadsheet classes
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

class grade_export_customexcel extends grade_export {

    public function __construct($course, $groupid = 0, $itemlist = null, $exportfeedback = false, $onlyactive = false) {
        parent::__construct($course, $groupid, $itemlist, $exportfeedback, $onlyactive);
    }

    public function print_grades() {
        global $CFG, $DB;

        $filename = clean_filename("grades-{$this->course->shortname}.xlsx");

        // Create PhpSpreadsheet workbook + worksheet.
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Results template sample');

        // Styles
        $headerStyle = [
            'font' => ['bold' => true],
            'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => 'F8CBAD'] // Orange, Accent 6, Lighter 60%
            ]
        ];
        $weightStyle = ['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]];
        $notesStyle = ['font' => ['italic' => true]];
        $notesBoldStyle = [
            'font' => ['bold' => true]];

        // -------------------------------
        // Metadata & Notes
        // -------------------------------
        $sheet->setCellValue('A1', 'Subject code');
        $sheet->setCellValue('B1', $this->course->shortname);

        $sheet->setCellValue('A2', 'Subject name');
        $sheet->setCellValue('B2', $this->course->fullname);

        $sheet->setCellValue('A3', 'Delivery mode');
        $sheet->setCellValue('B3', '---');

        $sheet->setCellValue('A4', 'Campus');
        $sheet->setCellValue('B4', '---');

        $sheet->setCellValue('D1', 'Please note:');
        $sheet->getStyle('D1')->applyFromArray($headerStyle);
        $sheet->setCellValue('D2', 'A dash (-) signifies a student that they did not submit the assessment and automatically fail the subject');
        $sheet->getStyle('D2')->applyFromArray($notesStyle);
        $sheet->setCellValue('D3', 'A zero (0) signifies a student has submitted an assessment but it was beyond the 2 week late assessment submission. They are still eligible to pass the subject if their overall total is greater than 50%.');
        $sheet->getStyle('D3')->applyFromArray($notesStyle);
        $sheet->setCellValue('D4', 'All Course totals are rounded to the whole number');
        $sheet->getStyle('D4')->applyFromArray($notesBoldStyle);

        // -------------------------------
        // Fetch items and students only
        // -------------------------------
        $items = grade_item::fetch_all(['courseid' => $this->course->id]);
        $context = context_course::instance($this->course->id);

        $studentrole = $DB->get_record('role', ['shortname' => 'student']);
        $users = [];
        if ($studentrole) {
            $users = get_role_users($studentrole->id, $context);
        }

        // Sort users by student number (idnumber) ascending
        usort($users, function($a, $b) {
            return strcmp($a->idnumber, $b->idnumber);
        });

        // -------------------------------
        // If no students → write message & exit
        // -------------------------------
        if (empty($users)) {
            $sheet->setCellValue('A6', 'No students are enrolled in this course or no grades available.');
            $sheet->getStyle('A6')->applyFromArray([
                'font' => ['bold' => true, 'italic' => true, 'color' => ['rgb' => 'FF0000']]
            ]);

            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header("Content-Disposition: attachment;filename=\"$filename\"");
            header('Cache-Control: max-age=0');

            $writer = new Xlsx($spreadsheet);
            $writer->save('php://output');
            exit;
        }

        // -------------------------------
        // Header row 1 (assessments, total, grade)
        // -------------------------------
        $row = 6;  // Start at row 6, row 5 stays blank
        $col = 4; // column D
        $assessmentitems = [];
        foreach ($items as $item) {
            if ($item->itemtype === 'mod') {
                $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
                $sheet->setCellValue($coord, $item->get_name());
                $sheet->getStyle($coord)->applyFromArray($headerStyle);
                $assessmentitems[] = $item;
                $col++;
            }
        }
        $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
        $sheet->setCellValue($coord, 'Total');
        $sheet->getStyle($coord)->applyFromArray($headerStyle);
        $col++;
        $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
        $sheet->setCellValue($coord, 'Grade');
        $sheet->getStyle($coord)->applyFromArray($headerStyle);

        // Enable text wrapping on row 6 (assessment names row)
        $sheet->getStyle('A6:' . \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . '6')
            ->getAlignment()
            ->setWrapText(true)
            ->setVertical(Alignment::VERTICAL_CENTER);

        // Optionally increase row height for visibility
        $sheet->getRowDimension(6)->setRowHeight(15);

        // -------------------------------
        // Header row 2 (StudentID etc + weightings)
        // -------------------------------
        $row++;
        $sheet->setCellValue('A' . $row, 'Student ID');
        $sheet->getStyle('A' . $row)->applyFromArray($headerStyle);
        $sheet->setCellValue('B' . $row, 'First name');
        $sheet->getStyle('B' . $row)->applyFromArray($headerStyle);
        $sheet->setCellValue('C' . $row, 'Surname');
        $sheet->getStyle('C' . $row)->applyFromArray($headerStyle);

        $col = 4;
        $totalweight = 0;
        foreach ($assessmentitems as $item) {
            $weight = $item->aggregationcoef;
            $totalweight += $weight;
            $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
            $sheet->setCellValue($coord, $weight . '%');
            $sheet->getStyle($coord)->applyFromArray($weightStyle);
            $col++;
        }
        $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
        $sheet->setCellValue($coord, $totalweight . '%');
        $sheet->getStyle($coord)->applyFromArray($weightStyle);
        $col++;
        $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
        $sheet->setCellValue($coord, '');
        $sheet->getStyle($coord)->applyFromArray($weightStyle);
        
        // -------------------------------
        // Set column widths
        // -------------------------------
        $sheet->getColumnDimension('A')->setWidth(14);
        $sheet->getColumnDimension('B')->setWidth(13);
        $sheet->getColumnDimension('C')->setWidth(13);

        // From D onward set width = 12
        for ($i = 4; $i <= $col; $i++) {
            $colLetter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($i);
            $sheet->getColumnDimension($colLetter)->setWidth(13);
        }

        // -------------------------------
        // Student rows
        // -------------------------------
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
                $sheet->setCellValue($coord, ($grade && $grade->finalgrade !== null) ? round($grade->finalgrade, 2) : '-');
                $col++;
            }

            if ($courseitem) {
                $coursegrade = grade_grade::fetch(['itemid' => $courseitem->id, 'userid' => $user->id]);
                $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
                $sheet->setCellValue($coord, ($coursegrade && $coursegrade->finalgrade !== null) ? round($coursegrade->finalgrade, 2) : '-');
                $col++;
                $coord = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($col) . $row;
                $sheet->setCellValue($coord, ($coursegrade && $coursegrade->finalgrade !== null) ? grade_format_gradevalue_letter($coursegrade->finalgrade, $courseitem) : '-');
            }
            $row++;
        }

        // -------------------------------
        // Output
        // -------------------------------
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header("Content-Disposition: attachment;filename=\"$filename\"");
        header('Cache-Control: max-age=0');

        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }
}
