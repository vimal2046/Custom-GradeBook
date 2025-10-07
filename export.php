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
require_once($CFG->dirroot . '/grade/export/grade_export_form.php');
require_once($CFG->dirroot . '/grade/export/customexcel/classes/exporter.php');

$id = required_param('id', PARAM_INT);

if (!$course = $DB->get_record('course', ['id' => $id])) {
    throw new moodle_exception('invalidcourseid');
}

require_login($course);
$context = context_course::instance($id);

require_capability('moodle/grade:export', $context);
require_capability('gradeexport/customexcel:view', $context);

// Build the same form used in index.php.
$formoptions = [
    'publishing' => true,
    'simpleui' => true,
    'multipledisplaytypes' => true,
];
$mform = new grade_export_form(null, $formoptions);

$formdata = $mform->get_data();

if (!$formdata) {
    // Support direct GET (publishing / dump style). Try to build formdata from URL params.
    $itemids = optional_param('itemids', '', PARAM_RAW);
    if ($itemids === '') {
        // No form data and no GET itemids — go back to index page.
        redirect(new moodle_url('/grade/export/customexcel/index.php', ['id' => $id]));
    }

    $exportfeedback = optional_param('export_feedback', 0, PARAM_BOOL);
    $onlyactive     = optional_param('export_onlyactive', 0, PARAM_BOOL);
    $displaytype    = optional_param('displaytype', $CFG->grade_export_displaytype, PARAM_RAW);
    $decimalpoints  = optional_param('decimalpoints', $CFG->grade_export_decimalpoints, PARAM_INT);

    // Use core helper to build a full $formdata object from raw GET values.
    $formdata = grade_export::export_bulk_export_data(
        $id,
        $itemids,
        $exportfeedback,
        $onlyactive,
        $displaytype,
        $decimalpoints
    );
} else {
    // Normal form POST path. Normalize display types and itemids.

    // 1) display types:
    if (!empty($formdata->display)) {
        if (is_array($formdata->display)) {
            // The form returns an associative array 'real' => GRADE_DISPLAY_TYPE_REAL, etc.
            $formdata->displaytype = $formdata->display;
        } else {
            // If it's a string for some reason, convert like core expects.
            $formdata->displaytype = grade_export::convert_flat_displaytypes_to_array($formdata->display);
        }
    }

    // 2) itemids: make a clean array of integer ids or -1 (all)
    if (!empty($formdata->itemids)) {
        if (is_array($formdata->itemids)) {
            // Form returns itemids as itemids[123] => 1. Keep keys that are truthy.
            $ids = [];
            foreach ($formdata->itemids as $key => $val) {
                if (!empty($val)) {
                    $ids[] = (int)$key;
                }
            }
            // If nothing selected keep -1 (consistent with core behaviour).
            $formdata->itemids = empty($ids) ? -1 : $ids;
        } else {
            // If string like "1,2,3" or "-1".
            if ($formdata->itemids === '-1') {
                $formdata->itemids = -1;
            } else {
                $parts = array_filter(array_map('trim', explode(',', $formdata->itemids)));
                $formdata->itemids = array_map('intval', $parts);
            }
        }
    } else {
        // No explicit itemids in form — default to -1 (all).
        $formdata->itemids = -1;
    }
}

// Create and run exporter using the full $formdata (this populates $this->displaytype, decimals, etc).
$export = new grade_export_customexcel($course, 0, $formdata);
$export->print_grades();
exit;
