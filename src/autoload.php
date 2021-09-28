<?php

spl_autoload_register(static function ($class) {
    $namespace = 'avadim\\FastExcelReader\\';
    if (0 === strpos($class, $namespace)) {
        include __DIR__ . '/FastExcelReader/' . str_replace($namespace, '', $class) . '.php';
    }
});

// EOF