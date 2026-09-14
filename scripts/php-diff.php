<?php

declare(strict_types=1);

/**
 * Cross-engine OPS parity helper: diff two schemas with the PHP holy-sheet and
 * replay the ops, so the TS port's Agent.diff / Agent.reduce can be compared
 * op for op. Uses the same minimal PSR-4 autoloader as php-describe.php (the
 * PHP core is zero-dependency) — no composer needed.
 *
 *   php php-diff.php <case.json>     {"a": ..., "b": ...}
 *                                    -> {"ops": Agent::diff(a, b), "reduced": Agent::reduce(a, ops)}
 *
 * The file may also hold a LIST of cases, answered with a list of results in the
 * same order, so a suite pays for one PHP process rather than one per case. A
 * case carrying "ops" instead of "b" replays those ops rather than diffing,
 * which pins the reducer on its own. A case {"hunks": [a, b]} answers
 * {"hunks": SheetDiff::hunks(a, b)}, which pins the alignment and its tie-break
 * directly. A case {"nested": {"leaf": x, "depth": n, "map": bool}} wraps x in
 * n arrays (lists, or maps under "k") HERE, and answers {"same": SheetDiff::same(v, v)},
 * which pins where json_encode's depth runs out without sending that nesting
 * through JSON. A case that throws answers {"error": ...}.
 *
 *   php php-diff.php --op-schema     -> Agent::opSchema()
 */

spl_autoload_register(function (string $class): void {
    $prefix = 'HolySheet\\';
    if (strncmp($class, $prefix, strlen($prefix)) !== 0) {
        return;
    }
    $rel = substr($class, strlen($prefix));
    // HOLY_SHEET_PHP_SRC first, sibling checkout second -- same as
    // php-tobytes.php and php-describe.php.
    $root = getenv('HOLY_SHEET_PHP_SRC') ?: __DIR__ . '/../../holy-sheet/src';
    $file = rtrim($root, '/') . '/' . str_replace('\\', '/', $rel) . '.php';
    if (is_file($file)) {
        require $file;
    }
});

$flags = JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE | JSON_THROW_ON_ERROR;

if (($argv[1] ?? null) === '--op-schema') {
    echo json_encode(\HolySheet\Agent::opSchema(), $flags);
    exit(0);
}

$input = json_decode((string) file_get_contents($argv[1]), true, 512, JSON_THROW_ON_ERROR);

$run = static function (array $case): array {
    try {
        if (array_key_exists('hunks', $case)) {
            return ['hunks' => \HolySheet\Ops\SheetDiff::hunks($case['hunks'][0], $case['hunks'][1])];
        }

        if (array_key_exists('nested', $case)) {
            $value = $case['nested']['leaf'];
            for ($i = 0; $i < $case['nested']['depth']; $i++) {
                $value = $case['nested']['map'] ? ['k' => $value] : [$value];
            }

            return ['same' => \HolySheet\Ops\SheetDiff::same($value, $value)];
        }

        $ops = array_key_exists('ops', $case)
            ? $case['ops']
            : \HolySheet\Agent::diff($case['a'], $case['b']);

        return ['ops' => $ops, 'reduced' => \HolySheet\Agent::reduce($case['a'], $ops)];
    } catch (\Throwable $e) {
        return ['error' => get_class($e) . ': ' . $e->getMessage()];
    }
};

echo json_encode(
    array_is_list($input) ? array_map($run, $input) : $run($input),
    $flags,
);
