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
 * Custom Excel grade export execution script.
 *
 * @package    gradeexport_customexcel
 * @copyright  2025 AC University
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */

require_once('../../../config.php');
require_once($CFG->dirroot . '/grade/export/lib.php');
require_once($CFG->dirroot . '/grade/export/customexcel/classes/exporter.php'); // Custom exporter class.

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

// Convert itemids array to string.
$itemlist = !empty($itemids) ? implode(',', $itemids) : '';

// Create exporter and trigger download.
$export = new grade_export_customexcel($course, $groupid, $itemlist, $exportfeedback, $onlyactive);
$export->print_grades();
exit;
