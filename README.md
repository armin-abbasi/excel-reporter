# excel-reporter
Create excel reports, from any data structure (collection, array, objects) with a line of code!

## Installation

```bash
composer require xerobase/excel-reporter
```

## Usage

Create an instance of Export class

```php
$exporter = new \Xerobase\ExcelReporter\Export();
```

You can simply export your data by calling export method :

```php
$books = \App\Models\Book::all();

$exporter->export($books);
```

Maybe want to filter some of unnecessary fields :

```php
$exporter->filterColumns(['id', 'created_at', 'updated_at'])->export($books);
```

Set direction to RTL :

```php
$exporter->setRightToLeft()->export($books);
```

Or change format to CSV :

```php
$exporter->setFormat('csv')->export($books);
```
