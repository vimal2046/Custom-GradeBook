<?php
require_once('../../../config.php');
require_once($CFG->dirroot.'/grade/export/lib.php');
require_once($CFG->dirroot.'/grade/export/customexcel/exporter.php'); // your class

$id             = required_param('id', PARAM_INT);
$groupid        = optional_param('groupid', 0, PARAM_INT);
$itemids        = optional_param_array('itemids', [], PARAM_INT);
$exportfeedback = optional_param('exportfeedback', 0, PARAM_BOOL);
$onlyactive     = optional_param('onlyactive', 0, PARAM_BOOL);

if (!$course = $DB->get_record('course', ['id' => $id])) {
    throw new moodle_exception('invalidcourseid');
}

require_login($course);
$context = context_course::instance($id);

require_capability('moodle/grade:export', $context);
require_capability('gradeexport/customexcel:view', $context);

// Convert itemids array → string
$itemlist = !empty($itemids) ? implode(',', $itemids) : '';

// Create exporter and trigger download
$export = new grade_export_customexcel($course, $groupid, $itemlist, $exportfeedback, $onlyactive);
$export->print_grades();
exit;
