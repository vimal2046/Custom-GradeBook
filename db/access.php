<?php
$capabilities = [
    'gradeexport/customexcel:view' => [
        'captype' => 'read',
        'contextlevel' => CONTEXT_COURSE,
        'archetypes' => [
            'teacher' => CAP_ALLOW,
            'editingteacher' => CAP_ALLOW,
            'manager' => CAP_ALLOW,
        ]
    ],
    'gradeexport/customexcel:view' => [
    'riskbitmask' => RISK_PERSONAL,
    'captype' => 'read',
    'contextlevel' => CONTEXT_COURSE,
    'archetypes' => ['teacher' => CAP_ALLOW, 'editingteacher' => CAP_ALLOW, 'manager' => CAP_ALLOW]
],

];
