<?php

declare(strict_types=1);

/**
 * Cross-engine parity helper: emit xlsx bytes from the PHP holy-sheet for a
 * given JSON schema, so the TS port can be diffed against it. Uses a minimal
 * PSR-4 autoloader (the PHP core is zero-dependency) — no composer needed.
 *
 *   php php-tobytes.php <schema.json> <out.xlsx>
 */

spl_autoload_register(function (string $class): void {
    $prefix = 'HolySheet\\';
    if (strncmp($class, $prefix, strlen($prefix)) !== 0) {
        return;
    }
    $rel = substr($class, strlen($prefix));
    // HOLY_SHEET_PHP_SRC first, sibling checkout second. The hard-coded
    // sibling path alone meant this only resolved inside the .agi envelope,
    // so CI -- or anyone with a different layout -- silently got no parity
    // run at all rather than an error.
    $root = getenv('HOLY_SHEET_PHP_SRC') ?: __DIR__ . '/../../holy-sheet/src';
    $file = rtrim($root, '/') . '/' . str_replace('\\', '/', $rel) . '.php';
    if (is_file($file)) {
        require $file;
    }
});

$schemaJson = file_get_contents($argv[1]);
$schema = json_decode($schemaJson, true);
$bytes = \HolySheet\Agent::toBytes($schema);
file_put_contents($argv[2], $bytes);
