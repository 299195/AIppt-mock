<?php
declare(strict_types=1);

/**
 * Local bridge for AIppt2:
 * php local_generate_pptx.php <template_json> <outline_md> <content_md> <output_pptx> [author] [last_page_text]
 */

error_reporting(E_ALL & ~E_DEPRECATED & ~E_STRICT & ~E_NOTICE & ~E_WARNING);

require_once(__DIR__ . '/AiToPPTX/include.inc.php');

function fail(string $message, int $code = 1): void
{
    fwrite(STDERR, $message . PHP_EOL);
    exit($code);
}

if ($argc < 5) {
    fail('Usage: php local_generate_pptx.php <template_json> <outline_md> <content_md> <output_pptx> [author] [last_page_text]');
}

$templatePath = (string)$argv[1];
$outlinePath = (string)$argv[2];
$contentPath = (string)$argv[3];
$outputPath = (string)$argv[4];
$author = isset($argv[5]) ? (string)$argv[5] : '';
$lastPageText = isset($argv[6]) ? (string)$argv[6] : 'Thank you';

if (!is_file($templatePath)) {
    fail("Template file not found: {$templatePath}");
}
if (!is_file($outlinePath)) {
    fail("Outline markdown file not found: {$outlinePath}");
}
if (!is_file($contentPath)) {
    fail("Content markdown file not found: {$contentPath}");
}

$outputDir = dirname($outputPath);
if (!is_dir($outputDir) && !mkdir($outputDir, 0777, true) && !is_dir($outputDir)) {
    fail("Failed to create output directory: {$outputDir}");
}

$templateRaw = file_get_contents($templatePath);
if ($templateRaw === false) {
    fail("Failed to read template file: {$templatePath}");
}
$templateJson = json_decode($templateRaw, true);
if (!is_array($templateJson)) {
    fail("Template JSON decode failed: {$templatePath}");
}

$outlineMarkdown = file_get_contents($outlinePath);
if ($outlineMarkdown === false) {
    fail("Failed to read outline markdown: {$outlinePath}");
}
$contentMarkdown = file_get_contents($contentPath);
if ($contentMarkdown === false) {
    fail("Failed to read content markdown: {$contentPath}");
}

$personalInfo = [
    'Author' => $author,
    'LastPageText' => $lastPageText,
];

try {
    $jsonData = Markdown_To_JsonData(
        $outlineMarkdown,
        $contentMarkdown,
        $templateJson,
        true,
        $personalInfo,
        0
    );

    $cacheRoot = __DIR__ . DIRECTORY_SEPARATOR . 'cache';
    if (!is_dir($cacheRoot) && !mkdir($cacheRoot, 0777, true) && !is_dir($cacheRoot)) {
        fail("Failed to create cache directory: {$cacheRoot}");
    }

    $targetCacheDir = $cacheRoot . DIRECTORY_SEPARATOR . 'local_' . date('Ymd_His') . '_' . mt_rand(1000, 9999);
    AiToPptx_MakePptx($jsonData, $targetCacheDir, $outputPath);
} catch (Throwable $e) {
    fail('Ai-To-PPTX export failed: ' . $e->getMessage(), 2);
}

if (!is_file($outputPath)) {
    fail('AiToPptx_MakePptx did not produce output.', 3);
}

echo $outputPath . PHP_EOL;
exit(0);
