<?php
require_once(__DIR__ . '/../../../config.php');

require_once('export.php');

// $courseid = required_param('id', PARAM_INT);
// $course = get_course($courseid);

// require_login($course);
// $context = context_course::instance($courseid);
// require_capability('gradeexport/customexcel:view', $context);

// $export = new grade_export_customexcel($courseid);
// $export->print_grades();


//-----------------------------------

// $courseid = required_param('id', PARAM_INT);
// $course = get_course($courseid);

// require_login($course);
// $context = context_course::instance($courseid);
// require_capability('gradeexport/customexcel:view', $context);

// // Pass the full course object, not just ID
// $export = new grade_export_customexcel($course);
// $export->print_grades();

require_once(__DIR__ . '/../../../config.php');
require_once($CFG->dirroot . '/grade/export/customexcel/export.php');

$courseid = required_param('id', PARAM_INT);
$course = get_course($courseid);

require_login($course);
$context = context_course::instance($courseid);
require_capability('gradeexport/customexcel:view', $context);

//  Pass the full course object, not just ID
$export = new grade_export_customexcel($course);
$export->print_grades();
