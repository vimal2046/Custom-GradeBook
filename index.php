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
 * Custom Excel grade export index page.
 *
 * @package    gradeexport_customexcel
 * @category   output
 * @copyright  2025 AC University
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */

require_once('../../../config.php');
require_once($CFG->dirroot . '/grade/export/lib.php');
require_once($CFG->dirroot . '/grade/export/grade_export_form.php');

$id = required_param('id', PARAM_INT); // Course ID.

$PAGE->set_url('/grade/export/customexcel/index.php', ['id' => $id]);

if (!$course = $DB->get_record('course', ['id' => $id])) {
    throw new moodle_exception('invalidcourseid');
}

require_login($course);
$context = context_course::instance($id);

require_capability('moodle/grade:export', $context);
require_capability('gradeexport/customexcel:view', $context);

// Setup page heading & action bar.
$actionbar = new \core_grades\output\export_action_bar($context, null, 'customexcel');
print_grade_page_head(
    $COURSE->id,
    'export',
    'customexcel',
    get_string('exportto', 'grades') . ' ' . get_string('pluginname', 'gradeexport_customexcel'),
    false,
    false,
    true,
    null,
    null,
    null,
    $actionbar
);

export_verify_grades($COURSE->id);

// Publishing support.
if (!empty($CFG->gradepublishing)) {
    $CFG->gradepublishing = has_capability('gradeexport/customexcel:publish', $context);
}

// Export form options.
$actionurl = new moodle_url('/grade/export/customexcel/export.php');
$formoptions = [
    'publishing' => true,
    'simpleui' => true,
    'multipledisplaytypes' => true,
];

// Create export form.
$mform = new grade_export_form($actionurl, $formoptions);

// Handle groups.
$groupmode    = groups_get_course_groupmode($course);
$currentgroup = groups_get_course_group($course, true);
if ($groupmode == SEPARATEGROUPS && !$currentgroup && !has_capability('moodle/site:accessallgroups', $context)) {
    echo $OUTPUT->heading(get_string('notingroup'));
    echo $OUTPUT->footer();
    die;
}

// Group selector.
groups_print_course_menu($course, 'index.php?id=' . $id);
echo '<div class="clearer"></div>';

// Display form (Download button etc).
$mform->display();

// Page footer.
echo $OUTPUT->footer();
