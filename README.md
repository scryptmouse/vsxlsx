# VSXLSX
A (Very Simple) XLSX Parser

## Installation

Install with [Composer](http://getcomposer.org).

```json
{
	"require": {
		"lazuli/vsxlsx": "0.0.1"
	}
}
```

## Usage

```php
require 'vendor/autoload.php';

$parser = new VSXLSX\Parser($filename, /*sheet number */);

if ($parser->parse()) {
	// Returns an array
	$parsed = $parser->get_parsed();
	/* Do something with the parsed array */
} else {
	foreach($parser->get_errors() as $error) {
		echo "Error: $error\n";
	}
}
```

## Column Names
### With a header
By default, the parser assumes there is a 'header' row. The first row will be used to generate column titles. They will be lowercased and have whitespace replaced by underscores.

Assuming cell A1 is _'Product Name'_, the associated key in the resultant array from `get_parsed()` will be `product_name`.

The column titles can be overridden with an array passed to the `header_names` method before parsing. This array can use numeric or alphabetical indices, interchangeably.

```php
	$column_names = array();

	// Any of these will override the first column's title
	$column_names[0] = 'product_title';
	$column_names['a'] = 'product_title';
	$column_names['A'] = 'product_title';

	// This will add a title for the 27th column (index 26)
	$column_names['aa'] = 'image_url';

	$parser->header_names($column_names);
```

### Without a header
If the parser is missing the header row, use the `has_header_row` method with `false` before parsing.

```php
	$parser->has_header_row(false);
```

## API Docs
Available [here](http://scryptmouse.github.com/vsxlsx/).

## License
MIT
