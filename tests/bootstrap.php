<?php

$vendorDir = __DIR__ . '/../../..';

if (file_exists($file = $vendorDir . '/autoload.php')) {
    require_once $file;
} else if (file_exists($file = './vendor/autoload.php')) {
    require_once $file;
} else {
    throw new \RuntimeException('Not found composer autoload');
}

// Test support classes (the package has no autoload-dev section)
foreach (glob(__DIR__ . '/Support/*.php') as $supportFile) {
    require_once $supportFile;
}
