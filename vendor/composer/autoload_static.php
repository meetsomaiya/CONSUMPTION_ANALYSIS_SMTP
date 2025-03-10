<?php

// autoload_static.php @generated by Composer

namespace Composer\Autoload;

class ComposerStaticInit299ae3db9247d48580081428a46fb170
{
    public static $prefixLengthsPsr4 = array (
        'P' => 
        array (
            'PHPMailer\\PHPMailer\\' => 20,
        ),
    );

    public static $prefixDirsPsr4 = array (
        'PHPMailer\\PHPMailer\\' => 
        array (
            0 => __DIR__ . '/..' . '/phpmailer/phpmailer/src',
        ),
    );

    public static $classMap = array (
        'Composer\\InstalledVersions' => __DIR__ . '/..' . '/composer/InstalledVersions.php',
    );

    public static function getInitializer(ClassLoader $loader)
    {
        return \Closure::bind(function () use ($loader) {
            $loader->prefixLengthsPsr4 = ComposerStaticInit299ae3db9247d48580081428a46fb170::$prefixLengthsPsr4;
            $loader->prefixDirsPsr4 = ComposerStaticInit299ae3db9247d48580081428a46fb170::$prefixDirsPsr4;
            $loader->classMap = ComposerStaticInit299ae3db9247d48580081428a46fb170::$classMap;

        }, null, ClassLoader::class);
    }
}
