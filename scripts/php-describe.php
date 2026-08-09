<?php

declare(strict_types=1);

/**
 * Cross-engine READER parity helper: read an .xlsx file with the PHP
 * holy-sheet reader and emit the resulting schema as JSON, so the TS port's
 * reader can be diffed against it. Uses the same minimal PSR-4 autoloader as
 * php-tobytes.php (the PHP core is zero-dependency) — no composer needed.
 *
 *   php php-describe.php <in.xlsx>
 */

spl_autoload_register(function (string $class): void {
    $prefix = 'HolySheet\\';
    if (strncmp($class, $prefix, strlen($prefix)) !== 0) {
        return;
    }
    $rel = substr($class, strlen($prefix));
    // HOLY_SHEET_PHP_SRC first, sibling checkout second -- same as
    // php-tobytes.php. This script was missed when that one was fixed, and CI
    // caught it the very first time it actually ran the parity suites.
    $root = getenv('HOLY_SHEET_PHP_SRC') ?: __DIR__ . '/../../holy-sheet/src';
    $file = rtrim($root, '/') . '/' . str_replace('\\', '/', $rel) . '.php';
    if (is_file($file)) {
        require $file;
    }
});

$path = $argv[1];
$result = \HolySheet\Agent::describe($path);
echo json_encode($result, JSON_PRETTY_PRINT);
