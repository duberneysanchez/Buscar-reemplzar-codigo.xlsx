<?php











namespace Composer;

use Composer\Semver\VersionParser;






class InstalledVersions
{
private static $installed = array (
  'root' => 
  array (
    'pretty_version' => 'dev-main',
    'version' => 'dev-main',
    'aliases' => 
    array (
    ),
    'reference' => '4a6303af59d8cd3699e71ee83ffb74e033c96e0a',
    'name' => '__root__',
  ),
  'versions' => 
  array (
    '__root__' => 
    array (
      'pretty_version' => 'dev-main',
      'version' => 'dev-main',
      'aliases' => 
      array (
      ),
      'reference' => '4a6303af59d8cd3699e71ee83ffb74e033c96e0a',
    ),
    'maennchen/zipstream-php' => 
    array (
      'pretty_version' => '3.1.0',
      'version' => '3.1.0.0',
      'aliases' => 
      array (
      ),
      'reference' => 'b8174494eda667f7d13876b4a7bfef0f62a7c0d1',
    ),
    'markbaker/complex' => 
    array (
      'pretty_version' => '3.0.2',
      'version' => '3.0.2.0',
      'aliases' => 
      array (
      ),
      'reference' => '95c56caa1cf5c766ad6d65b6344b807c1e8405b9',
    ),
    'markbaker/matrix' => 
    array (
      'pretty_version' => '3.0.1',
      'version' => '3.0.1.0',
      'aliases' => 
      array (
      ),
      'reference' => '728434227fe21be27ff6d86621a1b13107a2562c',
    ),
    'phpoffice/phpspreadsheet' => 
    array (
      'pretty_version' => '2.0.0',
      'version' => '2.0.0.0',
      'aliases' => 
      array (
      ),
      'reference' => '4a77798f835119754961a97714f135826a323caa',
    ),
    'psr/http-client' => 
    array (
      'pretty_version' => '1.0.3',
      'version' => '1.0.3.0',
      'aliases' => 
      array (
      ),
      'reference' => 'bb5906edc1c324c9a05aa0873d40117941e5fa90',
    ),
    'psr/http-factory' => 
    array (
      'pretty_version' => '1.0.2',
      'version' => '1.0.2.0',
      'aliases' => 
      array (
      ),
      'reference' => 'e616d01114759c4c489f93b099585439f795fe35',
    ),
    'psr/http-message' => 
    array (
      'pretty_version' => '2.0',
      'version' => '2.0.0.0',
      'aliases' => 
      array (
      ),
      'reference' => '402d35bcb92c70c026d1a6a9883f06b2ead23d71',
    ),
    'psr/simple-cache' => 
    array (
      'pretty_version' => '3.0.0',
      'version' => '3.0.0.0',
      'aliases' => 
      array (
      ),
      'reference' => '764e0b3939f5ca87cb904f570ef9be2d78a07865',
    ),
    'symfony/polyfill-iconv' => 
    array (
      'pretty_version' => 'v1.29.0',
      'version' => '1.29.0.0',
      'aliases' => 
      array (
      ),
      'reference' => 'cd4226d140ecd3d0f13d32ed0a4a095ffe871d2f',
    ),
    'symfony/polyfill-intl-grapheme' => 
    array (
      'pretty_version' => 'v1.29.0',
      'version' => '1.29.0.0',
      'aliases' => 
      array (
      ),
      'reference' => '32a9da87d7b3245e09ac426c83d334ae9f06f80f',
    ),
    'symfony/polyfill-intl-normalizer' => 
    array (
      'pretty_version' => 'v1.29.0',
      'version' => '1.29.0.0',
      'aliases' => 
      array (
      ),
      'reference' => 'bc45c394692b948b4d383a08d7753968bed9a83d',
    ),
    'symfony/polyfill-mbstring' => 
    array (
      'pretty_version' => 'v1.29.0',
      'version' => '1.29.0.0',
      'aliases' => 
      array (
      ),
      'reference' => '9773676c8a1bb1f8d4340a62efe641cf76eda7ec',
    ),
    'symfony/polyfill-php72' => 
    array (
      'pretty_version' => 'v1.29.0',
      'version' => '1.29.0.0',
      'aliases' => 
      array (
      ),
      'reference' => '861391a8da9a04cbad2d232ddd9e4893220d6e25',
    ),
    'voku/anti-xss' => 
    array (
      'pretty_version' => '4.1.42',
      'version' => '4.1.42.0',
      'aliases' => 
      array (
      ),
      'reference' => 'bca1f8607e55a3c5077483615cd93bd8f11bd675',
    ),
    'voku/portable-ascii' => 
    array (
      'pretty_version' => '2.0.1',
      'version' => '2.0.1.0',
      'aliases' => 
      array (
      ),
      'reference' => 'b56450eed252f6801410d810c8e1727224ae0743',
    ),
    'voku/portable-utf8' => 
    array (
      'pretty_version' => '6.0.13',
      'version' => '6.0.13.0',
      'aliases' => 
      array (
      ),
      'reference' => 'b8ce36bf26593e5c2e81b1850ef0ffb299d2043f',
    ),
  ),
);







public static function getInstalledPackages()
{
return array_keys(self::$installed['versions']);
}









public static function isInstalled($packageName)
{
return isset(self::$installed['versions'][$packageName]);
}














public static function satisfies(VersionParser $parser, $packageName, $constraint)
{
$constraint = $parser->parseConstraints($constraint);
$provided = $parser->parseConstraints(self::getVersionRanges($packageName));

return $provided->matches($constraint);
}










public static function getVersionRanges($packageName)
{
if (!isset(self::$installed['versions'][$packageName])) {
throw new \OutOfBoundsException('Package "' . $packageName . '" is not installed');
}

$ranges = array();
if (isset(self::$installed['versions'][$packageName]['pretty_version'])) {
$ranges[] = self::$installed['versions'][$packageName]['pretty_version'];
}
if (array_key_exists('aliases', self::$installed['versions'][$packageName])) {
$ranges = array_merge($ranges, self::$installed['versions'][$packageName]['aliases']);
}
if (array_key_exists('replaced', self::$installed['versions'][$packageName])) {
$ranges = array_merge($ranges, self::$installed['versions'][$packageName]['replaced']);
}
if (array_key_exists('provided', self::$installed['versions'][$packageName])) {
$ranges = array_merge($ranges, self::$installed['versions'][$packageName]['provided']);
}

return implode(' || ', $ranges);
}





public static function getVersion($packageName)
{
if (!isset(self::$installed['versions'][$packageName])) {
throw new \OutOfBoundsException('Package "' . $packageName . '" is not installed');
}

if (!isset(self::$installed['versions'][$packageName]['version'])) {
return null;
}

return self::$installed['versions'][$packageName]['version'];
}





public static function getPrettyVersion($packageName)
{
if (!isset(self::$installed['versions'][$packageName])) {
throw new \OutOfBoundsException('Package "' . $packageName . '" is not installed');
}

if (!isset(self::$installed['versions'][$packageName]['pretty_version'])) {
return null;
}

return self::$installed['versions'][$packageName]['pretty_version'];
}





public static function getReference($packageName)
{
if (!isset(self::$installed['versions'][$packageName])) {
throw new \OutOfBoundsException('Package "' . $packageName . '" is not installed');
}

if (!isset(self::$installed['versions'][$packageName]['reference'])) {
return null;
}

return self::$installed['versions'][$packageName]['reference'];
}





public static function getRootPackage()
{
return self::$installed['root'];
}







public static function getRawData()
{
return self::$installed;
}



















public static function reload($data)
{
self::$installed = $data;
}
}
