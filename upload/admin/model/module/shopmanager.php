<?php

ini_set('include_path', ini_get('include_path') . ':' . DIR_SYSTEM . 'library/Shopmanager/');
ini_set("max_execution_time", 360);

static $config = NULL;
static $log = NULL;

// Error Handlers
function shopmanager_error_handler($errno, $errstr, $errfile, $errline) {
	global $config;
	global $log;

	switch ($errno) {
		case E_NOTICE:
		case E_USER_NOTICE:
			$errors = "Notice";
			break;
		case E_WARNING:
		case E_USER_WARNING:
			$errors = "Warning";
			break;
		case E_ERROR:
		case E_USER_ERROR:
			$errors = "Fatal Error";
			break;
		default:
			$errors = "Unknown";
			break;
	}

	if (($errors == 'Warning') || ($errors == 'Unknown')) {
		return true;
	}

	if ($config->get('config_error_display')) {
		echo '<b>' . $errors . '</b>: ' . $errstr . ' in <b>' . $errfile . '</b> on line <b>' . $errline . '</b>';
	}

	if ($config->get('config_error_log')) {
		$log->write('PHP ' . $errors . ':  ' . $errstr . ' in ' . $errfile . ' on line ' . $errline);
	}

	return true;
}

function shopmanager_error_shutdown_handler() {
	$last_error = error_get_last();
	if ($last_error['type'] === E_ERROR) {
		shopmanager_error_handler(E_ERROR, $last_error['message'], $last_error['file'], $last_error['line']);
	}
}

class ModelModuleShopmanager extends Model {

    public $configTable = 'shopmanager_config';

	/**
	 * Current language
	 * @var type
	 */
	protected $languageId = 1;


     /**
	 * Cell's format
	 *
	 * @var type
	 */
	protected $priceFormat, $boxFormat, $weightFormat, $textFormat;
	protected $columns = array();
    public $settings = array();
    public $relation = array();
    protected $tablesMask = false;

	protected function init() {
		global $config;
		global $log;
		$config = $this->config;
		$log = $this->log;
		set_error_handler('shopmanager_error_handler', E_ALL);
		register_shutdown_function('shopmanager_error_shutdown_handler');
		$this->loadDefaultLanguage();
	}

    public function getConfig()
    {
        $sql = "SELECT * FROM " . DB_PREFIX . $this->configTable ;

        $result = $this->db->query($sql);

        $settings = array();
        if ($result->rows) {
            foreach ($result->rows as $row) {
                $settings[$row['group']][$row['field']] = $row['field'];
            }
        }
        $this->settings = $settings;
        return $settings;
    }

    public function setConfig($group, $data)
    {
        $fieldList = array();
        if ($group == 'product') {
            $fieldList = $this->getProductFildsList();
        }
        if ($group == 'category') {
            $fieldList = $this->getCateforyFildsList();
        }

        switch ($group) {
            case 'category':
            case 'product':
                $sql = "DELETE  FROM " . DB_PREFIX . $this->configTable . " WHERE `group` = '{$group}' ;";
                $this->db->query($sql);
                $sql = "INSERT INTO `" . DB_PREFIX . $this->configTable . "` VALUES ";
                $flag = false;
                foreach ($data as $key => $groupElement) {
                    if (isset($fieldList[$groupElement])) {
                        $sql .= "(NULL, '{$group}', '{$groupElement}'),";
                        $flag = true;
                    }
                }

                if ($flag) {
                    $sql = substr($sql,0,-1);
                    $this->db->query($sql .= ';');
                }
                break;
            default:
            break;
        }
    }


    /**
     *  Category fields list
     *
     */
    public function getCateforyFildsList()
    {
        return array(
            "category_id" => array(
                'name' => 'ID',
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'multirow' => false,
            ),
            "parent_id" => array(
                'name' => 'Parent category',
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'multirow' => false,
            ),
            "name" => array(
                'name' => 'Name',
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'multirow' => true,
            ),
            "sort_order" => array(
                'name' => 'Sort order',
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'multirow' => false,
            ),
            "image" => array(
                'name' => 'Image',
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'multirow' => false,
            ),
            "date_added" => array(
                'name' => 'Added date',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true,
                'multirow' => false,
            ),
            "date_modified" => array(
                'name' => 'Modify name',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true,
                'multirow' => false,
            ),
            "keyword" => array(
                'name' => 'Keyword',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true,
                'multirow' => false,
            ),
            "description" => array(
                'name' => 'Description',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true,
                'multirow' => true,
            ),
            "meta_description" => array(
                'name' => 'Meta description',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true,
                'multirow' => true,
            ),
            "meta_keyword" => array(
                'name' => 'Meta keyword',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true,
                'multirow' => true,
            ),
            "store_ids" => array(
                'name' => 'Stores',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true,
                'multirow' => false,
            ),
            "status" => array(
                'name' => 'Status',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true,
                'multirow' => false,
            )
        );
    }

    /**
     *  product fields list
     *
     */
    public function getProductFildsList()
    {
        return array(
            "product_id" => array(
                'name' => 'ID',
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'multirow' => false,
            ),
            "name" => array(
                'name' => 'Name',
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'multirow' => true,
            ),
            "categories" => array(
                'name' => 'Categories',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true,
                'multirow' => false,
            ),
            "sku" => array(
                'name' => 'SKU',
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'multirow' => false,
            ),
            "location" => array(
                'name' => 'Location',
                'length' => 20,
                'format' => null,
                'enabled' => false,
                'multirow' => false,
            ),
            "quantity" => array(
                'name' => 'Quantity',
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'multirow' => false,
            ),
            "model" => array(
                'name' => 'Model',
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'multirow' => false,
            ),
            "manufacturer" => array(
                'name' => 'Manufacturer',
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'multirow' => false,
            ),
            "image" => array(
                'name' => 'Image',
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'multirow' => false,
            ),
            "shipping" => array(
                'name' => 'Shiping',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false,
                'multirow' => false,
            ),
            "price" => array(
                'name' => 'Price',
                'length' => 20,
                'format' => $this->priceFormat,
                'enabled' => true,
                'multirow' => false,
            ),
            "weight" => array(
                'name' => 'weight',
                'length' => 20,
                'format' => $this->weightFormat,
                'enabled' => false,
                'multirow' => false,
            ),
            "unit" => array(
                'name' => 'Unit',
                'length' => 20,
                'format' => null,
                'enabled' => false,
                'multirow' => false,
            ),
            "length" => array(
                'name' => 'Length',
                'length' => 20,
                'format' => null,
                'enabled' => false,
                'multirow' => false,
            ),
            "width" => array(
                'name' => 'Width',
                'length' => 20,
                'format' => null,
                'enabled' => false,
                'multirow' => false,
            ),
            "height" => array(
                'name' => 'Height',
                'length' => 20,
                'format' => null,
                'enabled' => false,
                'multirow' => false,
            ),
            "length_unit" => array(
                'name' => 'Length unit',
                'length' => 20,
                'format' => null,
                'enabled' => false,
                'multirow' => false,
            ),
            "status" => array(
                'name' => 'Status',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true,
                'multirow' => false,
            ),
            "tax_class_id" => array(
                'name' => 'Tax class',
                'length' => 20,
                'format' => null,
                'enabled' => false,
                'multirow' => false,
            ),
            "viewed" => array(
                'name' => 'Viewed',
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'multirow' => false,
            ),
            "keyword" => array(
                'name' => 'Keyword',
                'length' => 20,
                'format' => null,
                'enabled' => false,
                'multirow' => false,
            ),
            "description" => array(
                'name' => 'Description',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false,
                'multirow' => true,
            ),
            "meta_description" => array(
                'name' => 'Meta description',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false,
                'multirow' => true,
            ),
            "meta_keyword" => array(
                'name' => 'Meta keyword',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false,
                'multirow' => true,
            ),
            "image" => array(
                'name' => 'Images',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false,
                'multirow' => false,
            ),
            "stock_status_id" => array(
                'name' => 'Stock status',
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'multirow' => false,
            ),
            "store_ids" => array(
                'name' => 'Store IDs',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false,
                'multirow' => false,
            ),
            "related" => array(
                'name' => 'Related',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false,
                'multirow' => false,
            ),
            "tags" => array(
                'name' => 'Tags',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false,
                'multirow' => true,
            ),
            "sort_order" => array(
                'name' => 'Sort order',
                'length' => 20,
                'format' => null,
                'enabled' => false,
                'multirow' => false,
            ),
            "subtract" => array(
                'name' => 'Subtract',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false,
                'multirow' => false,
            ),
            "minimum" => array(
                'name' => 'Minimum',
                'length' => 20,
                'format' => null,
                'enabled' => false,
                'multirow' => false,
            ),
            "date_added" => array(
                'name' => 'Date added',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true,
                'multirow' => false,
            ),
            "date_modified" => array(
                'name' => 'Date modified',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true,
                'multirow' => false,
            ),
            "date_available" => array(
                'name' => 'Date available',
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true,
                'multirow' => false,
            ),
        );
       /* return array(
            'product.product_id' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            'product_description.name' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            'product_to_category.categories' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            'product.sku' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            'product.location' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            'product.quantity' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            'product.model' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            'manufacturer.name' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            'product_image.image' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            'product.shipping' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            'product.price' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->priceFormat,
                'enabled' => true
            ),
            'product.weight' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->weightFormat,
                'enabled' => false
            ),
            'weight_class_description.unit' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            'product.length' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            'product.width' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            'product.height' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            'product.length_class_id' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            'product.status' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            'product.tax_class_id' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            'product.viewed' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            'product_description.language_id' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            'product_description.keyword' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            'product_description.description' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            'product_description.meta_description' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            'product_description.meta_keyword' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            'product_image.images' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            'product.stock_status_id' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            'product_to_store.store_id' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            'product_related.related' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            'product_description.tags' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            'product.sort_order' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            'product.subtract' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            'product.minimum' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            'product.date_added' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            'product.date_modified' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            'product.date_available' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
        );*/
    }

    public $tableSimple = array(
        'category' => array(
            'category_id' => 'NULL',
            'image' => '',
            'parent_id' => '0',
            'top' => '0',
            'column' => '1',
            'sort_order' => 0,
            'status' => '',
            'date_added' => 'NOW()',
            'date_modified' => 'NOW()',
        ),
        'category_description' => array(
            'category_id' => 'NULL',
            'language_id' => '1',
            'name' => '',
            'description' => '',
            'meta_description' => '',
            'meta_keyword' => '',
        ),
        'url_alias' => array(
            'url_alias_id' => 'NULL',
            'query' => '',
            'keyword' => '',
        ),
        'manufacturer' => array(
            'manufacturer_id' => 'NULL',
            'name' => '',
            'image' => 'NULL',
            'sort_order' => 0
        ),
        'product' => array(
            'product_id' => 'NULL',
            'model' => '',
            'sku' => '',
            'upc' => '',
            'ean' => '',
            'jan' => '',
            'isbn' => '',
            'mpn' => '',
            'location' => '',
            'quantity' => '1',
            'stock_status_id' => '8',
            'image' => '',
            'manufacturer_id' => '0',
            'shipping' => '1',
            'price' => '0',
            'points' => '0',
            'tax_class_id' => '0',
            'date_available' => 'NOW()',
            'weight' => '0',
            'weight_class_id' => '0',
            'length' => '0',
            'width' => '0',
            'height' => '0',
            'length_class_id' => '0',
            'subtract' => '1',
            'minimum' => '1',
            'sort_order' => '0',
            'status' => '1',
            'date_added' => 'NOW()',
            'date_modified' => 'NOW()',
            'viewed' => '0',
        ),
        'product_description' => array(
            'product_id' => '0',        //int(11) NOT NULL AUTO_INCREMENT,
            'language_id' => '0',       //int(11) NOT NULL,
            'name' => '',               //varchar(255) COLLATE utf8_bin NOT NULL,
            'description' => '',        //text COLLATE utf8_bin NOT NULL,
            'meta_description' => '',   //varchar(255) COLLATE utf8_bin NOT NULL,
            'meta_keyword' => '',       //varchar(255) COLLATE utf8_bin NOT NULL,
            'tag' => '',                //text COLLATE utf8_bin NOT NULL,
        ),
    );


    public function getSimple($tableName)
    {
        if (!isset($this->tablesMask[$tableName])) {
            if ( isset($this->tableSimple[$tableName])) {
                $sample = $this->tableSimple[$tableName];

            }
            else {
                return false;
            }
        }
        else {
            return $this->tablesMask[$tableName];
        }
    }

    /**
	 * Загрузка из XLS в MySQL
	 *
	 * @param type $filename
	 * @param type $group
	 * @return type
	 */
	public function upload($filename, $group)
    {
		$this->init();
		ini_set("memory_limit", "768M");

		require_once 'PHPExcel/Classes/PHPExcel.php';
		$cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_in_memory_serialized;
		$inputFileType = PHPExcel_IOFactory::identify($filename);
		$objReader = PHPExcel_IOFactory::createReader($inputFileType);
		$objReader->setReadDataOnly(true);
		PHPExcel_Settings::setCacheStorageMethod($cacheMethod);
		$reader = $objReader->load($filename);

		if (!$this->validateHeading($reader, $group)) {
			$this->log("Invalid {$group} header");
			$reader->disconnectWorksheets();
			unset($reader);
			return false;
		}

        switch (trim($group)) {
			case 'category':
				$this->cache->delete('category');
				$this->cache->delete('category_description');
				$this->cache->delete('url_alias');
				$result = $this->uploadCategories($reader);
				break;
			case 'product':
				$this->cache->delete('manufacturer');
				$this->cache->delete('product');
				$this->cache->delete('product_image');
				$this->cache->delete('product_option');
				$this->cache->delete('product_option_description');
				$this->cache->delete('product_option_value');
				$this->cache->delete('product_option_value_description');
				$this->cache->delete('product_to_category');
				$this->cache->delete('url_alias');
				$this->cache->delete('product_special');
				$this->cache->delete('product_discount');
				$result = $this->uploadProducts($reader);
				break;
			default:
				return false;
		}
		$reader->disconnectWorksheets();
		unset($reader);
		return $result;
	}

	/**
	 * Проверяет или поле должно выводиться
	 *
	 * @param type $column
	 * @return type
	 */
	protected function isEnabled($column) {
        if (!isset($this->columns[$column]['enabled'])) {
            return false;
        }
		return (bool) $this->columns[$column]['enabled'];
	}

	public function loadColumns($group) {
		switch ($group) {
			case 'category':
				$this->columns = $this->getCateforyFildsList();
				break;
			case 'product':
				$this->columns = $this->getProductFildsList();
				break;
			case 'product_options':
				break;
			case 'product_special':
				break;
			case 'product_discount':
				break;
			default:
				return false;
		}
		return true;
	}

	/**
	 * Получает сведения о загружаемых полях
	 *
	 * @param type $enabled
	 * @param type $keys
	 * @return type
	 */
	protected function getColumns($enabled = null, $keys = false) {
		if (empty($this->columns))
			return array();

		$columns = array();
		foreach ($this->columns as $key => $value) {
			if ((($enabled === true) && ($value['enabled'] !== true))
					|| (($enabled === false) && ($value['enabled'] !== false))) {
				continue;
			}
			$columns[$key] = $value;
		}

		return $keys ? array_keys($columns) : $columns;
	}

	/**
	 * Экспортирует выбранную группу данных
	 *
	 * @param type $group
	 * @return type
	 */
	public function export($group)
    {
		$this->init();
		ini_set("memory_limit", "128M");
		require_once 'Spreadsheet/Excel/Writer.php';

        // Creating a workbook
		$workbook = new Spreadsheet_Excel_Writer();
		$workbook->setVersion(8); // Use Excel97/2000 Format
		// Creating the categories worksheet

		$this->priceFormat = & $workbook->addFormat(array('Size' => 10, 'Align' => 'right', 'NumFormat' => '######0.00'));
		$this->boxFormat = & $workbook->addFormat(array('Size' => 10, 'vAlign' => 'vequal_space'));
		$this->weightFormat = & $workbook->addFormat(array('Size' => 10, 'Align' => 'right', 'NumFormat' => '##0.00'));
		$this->textFormat = & $workbook->addFormat(array('Size' => 10, 'NumFormat' => "@"));

		$group = strtolower(trim($group));
		// sending HTTP headers

        // TODO:remove test check after debugging
        if ($_REQUEST['test'] == 1) {
		    $workbook->send('backup_' . $group . '_' . date('Ymd') . '.xls');
        }
		$worksheetName = ucfirst(str_replace('_', ' ', $group));
		$worksheet = & $workbook->addWorksheet($worksheetName);

		$this->loadColumns($group);

		switch ($group) {
			case 'category':
                $this->columns = $this->getCateforyFildsList();
                $this->populateCategoriesWorksheet($worksheet);
				break;
			case 'product':
				$this->populateProductsWorksheet($worksheet);
				break;
			case 'product_options':
				$this->populateOptionsWorksheet($worksheet);
				break;
			case 'product_special':
				$this->populateSpecialsWorksheet($worksheet);
				break;
			case 'product_discount':
				$this->populateDiscountsWorksheet($worksheet);
				break;
			default:
				return false;
		}

		//$worksheet->freezePanes(array(1, 1, 1, 1));

		// Let's send the file
		$workbook->close();

		exit;
	}

	/**
	 * Очистка импортируемых данных
	 *
	 * @param type $str
	 * @param type $allowBlanks
	 * @return type
	 */
	protected function clean($str, $allowBlanks=false) {
		$result = "";
		$n = strlen($str);
		for ($m = 0; $m < $n; $m++) {
			$ch = substr($str, $m, 1);
			if (($ch == " ") && (!$allowBlanks) || ($ch == "\n") || ($ch == "\r") || ($ch == "\t") || ($ch == "\0") || ($ch == "\x0B")) {
				continue;
			}
			$result .= $ch;
		}
		return $result;
	}

	protected function import($sql) {
		foreach (explode(";\n", $sql) as $sql) {
			$sql = trim($sql);
			if ($sql) {
				$this->db->query($sql);
			}
		}
	}

	protected function getDefaultWeightUnit() {
		return $this->config->get('config_weight_class');
	}

	protected function getDefaultMeasurementUnit() {
		return $this->config->get('config_length_class');
	}

	protected function updateManufacturersInDB(&$products, &$manufacturerIds)
    {
		// find all manufacturers already stored in the database
        $keyArray = array_keys($products);
        $manufacturerName = array();
        if (isset($products[$keyArray[0]]['manufacturer'])) {
            foreach ($products as $product) {
                if ( trim($product['manufacturer']) != '') {
                    $manufacturerName[$product['manufacturer']] = $product['manufacturer'];
                }
            }
            $sql = "SELECT * FROM `" . DB_PREFIX . "manufacturer`;";
            $result = $this->db->query($sql);

            $query = '';
            foreach($manufacturerName as $mName) {
                $flag = false;
                foreach ($result->rows as $manufacturer) {
                    if ($manufacturer['name'] == $mName) {
                        $flag = true;
                        break;
                    }
                }

                if (!$flag) {
                    $manufacturerSimple = $this->tableSimple['manufacturer'];
                    $manufacturerSimple['name'] = $mName;

                    $query .= "('" .implode($manufacturerSimple, "','") . "')";
                    $query = str_replace("'NULL'",'NULL',$query);

                    $sql = "INSERT INTO `". DB_PREFIX . "manufacturer` VALUES {$query}";
                    $this->db->query($sql);

                    $lastId = $this->db->getLastId();

                    $this->load->model('setting/store');
                    $results = $this->model_setting_store->getStores();

                    $stores = array(0 => true);
                    foreach ($results as $store) {
                        $stores[$store['store_id']] = true;
                    }

                    $query = '';
                    foreach ($stores as $store_id => $store) {
                        $query .= "('" .implode(array( $lastId, $store_id ), "','") . "'),";
                    }

                    $query = rtrim(trim($query), ',');
                    $sql = "INSERT INTO `" . DB_PREFIX . "manufacturer_to_store` (`manufacturer_id`,`store_id`) VALUES {$query} ;";
                    $this->db->query($sql);
                }
            }
        }
        else {
            return true;
        }
		return true;
	}

	protected function getWeightClassIds() {
		// find all weight classes already stored in the database
		$weightClassIds = array();
		$sql = "SELECT `weight_class_id`, `unit` FROM `" . DB_PREFIX . "weight_class_description` WHERE `language_id`='{$this->languageId}';";
		$result = $this->db->query($sql);
		if ($result->rows) {
			foreach ($result->rows as $row) {
				$weightClassId = $row['weight_class_id'];
				$unit = $row['unit'];
				if (!isset($weightClassIds[$unit])) {
					$weightClassIds[$unit] = $weightClassId;
				}
			}
		}
		return $weightClassIds;
	}

	protected function getLengthClassIds() {
		// find all length classes already stored in the database
		$lengthClassIds = array();
		$sql = "SELECT `length_class_id`, `unit` FROM `" . DB_PREFIX . "length_class_description` WHERE `language_id`='{$this->languageId}';";
		$result = $this->db->query($sql);
		if ($result->rows) {
			foreach ($result->rows as $row) {
				$lengthClassId = $row['length_class_id'];
				$unit = $row['unit'];
				if (!isset($lengthClassIds[$unit])) {
					$lengthClassIds[$unit] = $lengthClassId;
				}
			}
		}
		return $lengthClassIds;
	}

    protected function getManufacturerByName( $name )
    {
        $defaultManufacturer = 0;

        $result = $this->db->query("SELECT `manufacturer_id` FROM `". DB_PREFIX ."manufacturer` WHERE name = '{$name}'; ");
        $data = $result->row;

        if ( isset($data['manufacturer_id']) ) {
            return $data['manufacturer_id'];
        }
        else {
            return $defaultManufacturer;
        }
    }


	protected function updateProductsInDB($products) {
		$this->import("START TRANSACTION;\n");

		// store or update manufacturers
		$manufacturerIds = array();
		if (!$this->updateManufacturersInDB($products, $manufacturerIds)) {
			$this->db->query('ROLLBACK;');
			return false;
        }

		// get weight classes
		$weightClassIds = $this->getWeightClassIds();

		// get length classes
		$lengthClassIds = $this->getLengthClassIds();

		// generate and execute SQL for storing the products

        foreach ($products as $_key => $product)
        {
            $productData = array();
            if (trim($product['product_id']) != '') {
                $sql = "SELECT * FROM `" . DB_PREFIX . "product` WHERE product_id = '{$product['product_id']}'; ";
                $result = $this->db->query($sql);
                $productData = $result->row;
            }

            $productData = array_merge($this->tableSimple['product'], $productData);
            $insert = '';

            if (isset($product['model'])) {
                $productData['model'] = $product['model'];
                $insert .= "`model` = '{$product['model']}',";
            }
            if (isset($product['sku'])) {
                $productData['sku'] = $product['sku'];
                $insert .= "`sku` = '{$product['sku']}',";
            }
            if (isset($product['location'])) {
                $productData['location'] = $product['location'];
                $insert .= "`location` = '{$product['location']}',";
            }
            if (isset($product['quantity'])) {
                $productData['quantity'] = $product['quantity'];
                $insert .= "`quantity` = '{$product['quantity']}',";
            }

            if (isset($product['image'])) {
                $img = explode(',', $product['image']);
                $productData['image'] = $img[0];
                $insert .= "`image` = '{$img[0]}',";
            }

            if (isset($product['manufacturer'])) {
                $manId = $this->getManufacturerByName($product['manufacturer']);
                $productData['manufacturer_id'] = $manId;
                $insert .= "`manufacturer_id` = '{$manId}',";
            }
            if (isset($product['shipping'])) {
                $productData['shipping'] = $product['shipping'];
                $insert .= "`shipping` = '{$product['shipping']}',";
            }
            if (isset($product['price'])) {
                $productData['price'] = $product['price'];
                $insert .= "price = '{$product['price']}',";
            }
            if (isset($product['tax_class_id'])) {
                $productData['tax_class_id'] = $product['tax_class_id'];
                $insert .= "tax_class_id = '{$product['tax_class_id']}',";
            }
            if (isset($product['date_available'])) {
                $productData['date_available'] = $product['date_available'];
                $insert .= "date_available = '{$product['date_available']}',";
            }
            if (isset($product['weight'])) {
                $productData['weight'] = $product['weight'];
                $insert .= "weight = '{$product['weight']}',";
            }
            if (isset($product['unit'])) {
                if ( isset($weightClassIds[$product['weight']])) {
                    $productData['weight'] = $weightClassIds[$product['weight']];
                    $insert .= "weight = '{$product['weight']}',";
                }
            }
            if (isset($product['length'])) {
                $productData['length'] = $product['length'];
                $insert .= "length = '{$product['length']}',";
            }
            if (isset($product['width'])) {
                $productData['width'] = $product['width'];
                $insert .= "width = '{$product['width']}',";
            }
            if (isset($product['height'])) {
                $productData['height'] = $product['height'];
                $insert .= "height = '{$product['height']}',";
            }
            if (isset($product['length_unit'])) {
                if ( isset($weightClassIds[$product['length_unit']])) {
                    $productData['length_unit'] = $lengthClassIds[$product['length_unit']];
                    $insert .= "length_unit = '{$product['length_unit']}',";
                }
            }
            if (isset($product['subtract'])) {
                $productData['subtract'] = $product['subtract'];
                $insert .= "subtract = '{$product['subtract']}',";
            }
            if (isset($product['minimum'])) {
                $productData['minimum'] = $product['minimum'];
                $insert .= "minimum = '{$product['minimum']}',";
            }
            if (isset($product['sort_order'])) {
                $productData['sort_order'] = $product['sort_order'];
                $insert .= "sort_order = '{$product['sort_order']}',";
            }
            if (isset($product['status'])) {
                $productData['status'] = $product['status'];
                $insert .= "status = '{$product['status']}',";
            }
            if (isset($product['date_added'])) {
                $productData['date_added'] = $product['date_added'];
                $insert .= "date_added = '{$product['date_added']}',";
            }
            if (isset($product['date_modified'])) {
                $productData['date_modified'] = $product['date_modified'];
                $insert .= "date_modified = '{$product['date_modified']}',";
            }
            if (isset($product['viewed'])) {
                $productData['viewed'] = $product['viewed'];
                $insert .= "viewed = '{$product['viewed']}',";
            }

            $query = "('".implode($productData, "','"). "')";
            $query = str_replace("'NULL'", 'NULL', $query);

            $insert = rtrim(trim($insert), ',');
            $sql = "INSERT INTO `" . DB_PREFIX . "product` VALUES {$query} ON DUPLICATE KEY UPDATE {$insert} ; ";
            $this->db->query($sql);

            $lastId = trim($product['product_id']);
            if ( $lastId == '') {
                $lastId = $this->db->getLastId();
            }

            //update product_to_category
            if (isset($product['categories'])) {
                $sql = "DELETE FROM `" . DB_PREFIX . "product_to_category` WHERE product_id = '{$lastId}'; ";
                $this->db->query($sql);
                $insert = '';
                $categories = explode(',', $product['categories']);
                foreach ($categories as $cat) {
                    $insert .= "('{$lastId}','{$cat}'),";
                }
                $insert = rtrim(trim($insert), ',');

                $sql = "INSERT INTO `" . DB_PREFIX . "product_to_category` VALUES {$insert} ; ";
                $this->db->query($sql);
            }

            //update product_to_store
            if (isset($product['store_ids'])) {
                $sql = "DELETE FROM `" . DB_PREFIX . "product_to_store` WHERE product_id = '{$lastId}'; ";
                $this->db->query($sql);
                $insert = '';
                $stores = explode(',', $product['store_ids']);
                foreach ($stores as $cat) {
                    $insert .= "('{$lastId}','{$cat}'),";
                }
                $insert = rtrim(trim($insert), ',');

                $sql = "INSERT INTO `" . DB_PREFIX . "product_to_store` VALUES {$insert} ; ";
                $this->db->query($sql);
            }

            //update product_to_image
            if (isset($product['image'])) {
                $sql = "DELETE FROM `" . DB_PREFIX . "product_image` WHERE product_id = '{$lastId}'; ";
                $this->db->query($sql);
                $insert = '';
                $images = explode(',', $product['image']);
                foreach ($images as $image) {
                    $insert .= "(NULL,'{$lastId}','{$image}', 0),";
                }
                $insert = rtrim(trim($insert), ',');

                $sql = "INSERT INTO `" . DB_PREFIX . "product_image` VALUES {$insert} ; ";
                $this->db->query($sql);
            }

            //UPDATE url_alias
            if (isset($product['keyword']) && trim($product['keyword']) != '') {
                $sql = "DELETE FROM `" . DB_PREFIX . "url_alias` WHERE query = 'product_id={$lastId}'; ";
                $this->db->query($sql);

                $sql = "INSERT INTO `" . DB_PREFIX . "url_alias` VALUES ('0','product_id={$lastId}','{$product['keyword']}');";
                $this->db->query($sql);
            }

            //UPDATE product_related
            if (isset($product['related'])) {
                $sql = "DELETE FROM `" . DB_PREFIX . "product_related` WHERE product_id = '{$lastId}'; ";
                $this->db->query($sql);
                $insert = '';
                $related = explode(',', $product['related']);
                foreach ($related as $rel) {
                    $insert .= "('{$lastId}','{$rel}'),";
                }
                $insert = rtrim(trim($insert), ',');

                $sql = "INSERT INTO `" . DB_PREFIX . "product_related` VALUES {$insert} ; ";
                $this->db->query($sql);
            }

            //UPDATE product_description
            $this->load->model('localisation/language');
            $languageList = $this->model_localisation_language->getLanguages();

            //if ( isset($product['name']) || isset($product['description']) || isset($product['meta_description']) || isset($product['meta_keyword']) ) {}

            foreach ($languageList as $code =>$language) {
                $sql = "SELECT * FROM `" . DB_PREFIX . "product_description` WHERE product_id = '{$lastId}' AND language_id = '{$language['language_id']}' ;";
                $result = $this->db->query($sql);

                $productDescription = $this->tableSimple['product_description'];
                $productDescription = array_merge($productDescription, $result->row);

                $flag = false;

                $productDescription['product_id'] = $lastId;
                $productDescription['language_id'] = $language['language_id'];

                if ( isset($product['name']) && isset($product['name'][$language['language_id']]) ) {
                    $productDescription['name'] = $product['name'][$language['language_id']];
                    $flag = true;
                }
                if ( isset($product['description']) && isset($product['description'][$language['language_id']]) ) {
                    $productDescription['description'] = $product['description'][$language['language_id']];
                    $flag = true;
                }
                if ( isset($product['meta_description']) && isset($product['meta_description'][$language['language_id']]) ) {
                    $productDescription['meta_description'] = $product['meta_description'][$language['language_id']];
                    $flag = true;
                }
                if ( isset($product['meta_keyword']) && isset($product['meta_keyword'][$language['language_id']]) ) {
                    $productDescription['meta_keyword'] = $product['meta_keyword'][$language['language_id']];
                    $flag = true;
                }

                if ($flag) {
                    $sql = "DELETE FROM `" . DB_PREFIX . "product_description` WHERE product_id = '{$lastId}' AND language_id = '{$language['language_id']}' ;";
                    $this->db->query($sql);

                    $query = "('".implode($productDescription, "','"). "')";
                    $query = str_replace("'NULL'", 'NULL', $query);

                    $sql = "INSERT INTO `" . DB_PREFIX . "product_description` VALUES {$query} ;";

                    $this->db->query($sql);
                }
            }
        }

		// final commit
		$this->db->query("COMMIT;");
		return true;
	}

	/**
	 * Определение кодировки строки
	 *
	 * @param type $str
	 * @return type
	 */
	protected function detect_encoding($str) {
		return mb_detect_encoding($str, 'UTF-8,ISO-8859-15,ISO-8859-1,cp1251,KOI8-R');
	}

	function uploadProducts(&$reader) {

        $data = $reader->getSheet(0);
        if ($data->getTitle() != 'Product') {
            $this->log('Wrong file type');
            return false;
        }

        $this->columns = $this->getProductFildsList();
        $this->buildFieldRelations($reader, 'product');
        $this->columns = $this->relation;

		$defaultWeightUnit = $this->getDefaultWeightUnit();
		$defaultMeasurementUnit = $this->getDefaultMeasurementUnit();
		$defaultStockStatusId = $this->config->get('config_stock_status_id');

		$products = array();
		$isFirstRow = true;
		$i = 0;
		$j = 1;
		$k = $data->getHighestRow();


		for ($i = 2; $i < $k; $i+=1) {
			$product = array();
			$productId = trim($this->getCell($data, $i, $j++));

            foreach ($this->relation as $relKey => $relation) {
                if ( isset($relation['position']) ) {
                    $tmp = trim($this->getCell($data, $i,$relation['position']));
                    $product[$relKey] = htmlentities($tmp, ENT_QUOTES, $this->detect_encoding($tmp));
                    //$product[$relKey] = $tmp;
                }
                else {
                    foreach ($relation as $langCode => $sublanguage) {
                        $tmp = trim($this->getCell($data, $i,$relation[$langCode]['position']));
                        $product[$relKey][$langCode] = htmlentities($tmp, ENT_QUOTES, $this->detect_encoding($tmp));
                        //$product[$relKey][$langCode] = $tmp;
                    }
                }
            }

			$products[] = $product;
			$j = 1;
        }
		return $this->updateProductsInDB($products);
	}

	/**
	 * Обновляет записи в БД
	 *
	 * @param type $categories
	 * @return type
	 */
	protected function updateCategoriesInDB(&$categories)
    {
		$this->import("START TRANSACTION;");
		// generate and execute SQL for inserting the categories

		foreach ($categories as $category) {
            $simpleCategory = $this->tableSimple['category'];

            //insert category data
            $insert = '';
            $isNew = true;
            if (isset($category['category_id']) && trim($category['category_id']) != '' ) {
                $sql = "SELECT * FROM ". DB_PREFIX ."category WHERE category_id = {$category['category_id']}";
                $result = $this->db->query($sql);
                $simpleCategory = array_merge($simpleCategory, $result->row);

                $simpleCategory['category_id'] = $category['category_id'];
                $isNew = false;
            }

            if (isset($category['image'])) {
                $simpleCategory['image'] = $category['image'];
                $insert .= "image = '{$category['image']}', ";
            }
            if (isset($category['parent_id'])) {
                $simpleCategory['parent_id'] = $category['parent_id'];
                $insert .= "parent_id = '{$category['parent_id']}', ";
            }
            if (isset($category['top'])) {
                $simpleCategory['top'] = $category['top'];
                $insert .= "top = '{$category['top']}', ";
            }
            if (isset($category['column'])) {
                $simpleCategory['column'] = $category['column'];
                $insert .= "column = '{$category['column']}', ";
            }
            if (isset($category['sort_order'])) {
                $simpleCategory['sort_order'] = $category['sort_order'];
                $insert .= "sort_order = '{$category['sort_order']}', ";
            }
            if (isset($category['status'])) {
                $simpleCategory['status'] = $category['status'];
                $insert .= "status = '{$category['status']}', ";
            }
            if (isset($category['date_added'])) {
                $simpleCategory['date_added'] = $category['date_added'];
                $insert .= "date_added = '{$category['date_added']}', ";
            }
            if (isset($category['date_modified'])) {
                $simpleCategory['date_modified'] = $category['date_modified'];
                $insert .= "date_modified = '{$category['date_modified']}', ";
            }

            $sql = "( '" . implode($simpleCategory, "','") . "' )";
            $sql = str_replace("'NULL'", "NULL", $sql);
            $insert = rtrim(trim($insert), ',');

            $query = "INSERT INTO `". DB_PREFIX ."category` VALUES {$sql} ON DUPLICATE KEY UPDATE {$insert} ";
            $this->db->query($query);

            $lastId = $this->db->getLastId();
            if (!$isNew) {
                $lastId = $category['category_id'];
            }

            //insert category_description
            $this->load->model('localisation/language');
            $languageList = $this->model_localisation_language->getLanguages();

            $catDesc = array();
            foreach ($languageList as $code => $language) {
                $catDescSimple = $this->tableSimple['category_description'];

                $sql = "SELECT * FROM `". DB_PREFIX ."category_description` WHERE category_id = {$lastId} and language_id = {$language['language_id']} LIMIT 1; ";
                $result = $this->db->query($sql);
                $catDescSimple = array_merge($catDescSimple, $result->row);
                $catDescSimple['category_id'] = $lastId;
                $catDescSimple['language_id'] = $language['language_id'];

                if (isset($category['name']) && isset($category['name'][$language['language_id']])) {
                    $catDescSimple['name'] = $category['name'][$language['language_id']];
                }

                if (isset($category['description']) && isset($category['description'][$language['language_id']])) {
                    $catDescSimple['description'] = $category['description'][$language['language_id']];
                }

                if (isset($category['meta_description']) && isset($category['meta_description'][$language['language_id']])) {
                    $catDescSimple['meta_description'] = $category['meta_description'][$language['language_id']];
                }

                if (isset($category['meta_keyword']) && isset($category['meta_keyword'][$language['language_id']])) {
                    $catDescSimple['meta_keyword'] = $category['meta_keyword'][$language['language_id']];
                }

                $sql = "( '" . implode($catDescSimple, "','") . "' )";
                $sql = str_replace("'NULL'", "NULL", $sql);

                $query = "DELETE FROM `". DB_PREFIX ."category_description` WHERE category_id = '{$catDescSimple['category_id']}' AND language_id = '{$catDescSimple['language_id']}';";
                $this->db->query($query);

                $query = "INSERT INTO `". DB_PREFIX ."category_description` VALUES {$sql} ";
                $this->db->query($query);
            }

            //adding keywords to  url_alias;
            if ( isset($category['keyword']) && trim($category['keyword']) != '' ) {
                $sql = "DELETE FROM " . DB_PREFIX . "url_alias  WHERE query = 'category_id={$lastId}'";
                $this->db->query($sql);

                $sql = "INSERT INTO `" . DB_PREFIX . "url_alias` (`query`,`keyword`) ";
                $sql .= "VALUES ('category_id={$lastId}','{$category['keyword']}') ";
                $this->db->query($sql);
            }

			/*if (isset($category['store_ids'])) {
				foreach ($category['store_ids'] as $storeId) {
					if ($storeId == '')
						continue;
                    $sql = "INSERT INTO `" . DB_PREFIX . "category_to_store` (`category_id`,`store_id`) ";
					$sql .= "VALUES ({$categoryId},{$storeId}) ";
					$sql .= "ON DUPLICATE KEY UPDATE `category_id`='{$categoryId}',`store_id`='{$storeId}';";
					$this->db->query($sql);
				}
			}*/
		}
		// final commit
		$this->db->query("COMMIT;");
		return true;
	}

    protected function buildFieldRelations(&$reader, $type)
    {
        $this->load->model('localisation/language');
        $languageList = $this->model_localisation_language->getLanguages();

        $this->relation = array();
        $data = $reader->getSheet(0);
        $relation = array();
        switch($type) {
            case 'category':
                $config = $this->getCateforyFildsList();
                $colCount = PHPExcel_Cell::columnIndexFromString($data->getHighestColumn());

                for ($i = 1; $i < $colCount+1; $i++ ) {
                    $fieldName = $this->getCell($data, 0, $i);
                    foreach ($config as $key => $configField) {
                        if ($fieldName == $key) {
                            $this->relation[$fieldName] = $config[$fieldName];
                            $this->relation[$fieldName]['position'] = $i;
                            break;
                        }
                        else if (preg_match("/^{$key}[_a-z]{0,3}$/", $fieldName)) {
                            $langArray = preg_split("/^{$key}[_]{1}/", $fieldName);
                            $languageCode = $languageList[$langArray[1]]['language_id'];
                            $this->relation[$key][$languageCode] = $config[$key];
                            $this->relation[$key][$languageCode]['position'] = $i;
                            break;
                        }
                    }
                }
            break;
            case 'product':
                $config = $this->getProductFildsList();
                $colCount = PHPExcel_Cell::columnIndexFromString($data->getHighestColumn());

                for ($i = 1; $i < $colCount+1; $i++ ) {
                    $fieldName = $this->getCell($data, 0, $i);
                    foreach ($config as $key => $configField) {
                        if ($fieldName == $key) {
                            $this->relation[$fieldName] = $config[$fieldName];
                            $this->relation[$fieldName]['position'] = $i;
                            break;
                        }
                        else if (preg_match("/^{$key}[_a-z]{0,3}$/", $fieldName)) {
                            $langArray = preg_split("/^{$key}[_]{1}/", $fieldName);
                            $languageCode = $languageList[$langArray[1]]['language_id'];
                            $this->relation[$key][$languageCode] = $config[$key];
                            $this->relation[$key][$languageCode]['position'] = $i;
                            break;
                        }
                    }
                }
            break;
            default:
                return false;
            break;
        }
    }

	protected function uploadCategories(&$reader) {
        $data = $reader->getSheet(0);

        if ($data->getTitle() != 'Category') {
            $this->log('Wrong file type');
            return false;
        }

        $this->columns = $this->getCateforyFildsList();
        $this->buildFieldRelations($reader, 'category');

		$categories = array();
		$isFirstRow = true;
		$i = 0;
		$j = 1;
		$k = $data->getHighestRow();

		for ($i = 1; $i < $k; $i+=1) {
			if ($isFirstRow) {
				$isFirstRow = false;
				continue;
			}
			$category = array();
            if (!isset($this->relation['category_id'])) {
                return;
            }

            foreach ($this->relation as $relKey => $relation) {
                if ( isset($relation['position']) ) {
                    $tmp = trim($this->getCell($data, $i,$relation['position']));
                    //$category[$relKey] = htmlentities($tmp, ENT_QUOTES, $this->detect_encoding($tmp));
                    $category[$relKey] = $tmp;
                }
                else {
                    foreach ($relation as $langCode => $sublanguage) {
                        $tmp = trim($this->getCell($data, $i,$relation[$langCode]['position']));
                        $category[$relKey][$langCode] = htmlentities($tmp, ENT_QUOTES, $this->detect_encoding($tmp));
                        $category[$relKey][$langCode] = $tmp;
                    }
                }
            }


			//    $categories[$category['category_id']] = $category;

            $categories[] = $category;
			$j = 1;
		}

		return $this->updateCategoriesInDB($categories);
	}

	function storeOptionNamesIntoDatabase(&$options, &$optionIds)
    {
// add option names, ids, and sort orders to the database
		$maxOptionId = 0;
		$sortOrder = 0;
		$sql = "INSERT INTO `" . DB_PREFIX . "product_option` (`product_option_id`, `product_id`, `sort_order`) VALUES ";
		$sql2 = "INSERT INTO `" . DB_PREFIX . "product_option_description` (`product_option_id`, `product_id`, `language_id`, `name`) VALUES ";
		$k = strlen($sql);
		$first = true;
		foreach ($options as $option) {
			$productId = $option['product_id'];
			$name = $option['option'];
			$langId = $option['language_id'];
			if ($productId == "") {
				continue;
			}

			if ($name == "") {
				continue;
			}
			if (!isset($optionIds[$productId][$name])) {
				$maxOptionId += 1;
				$optionId = $maxOptionId;
				if (!isset($optionIds[$productId])) {
					$optionIds[$productId] = array();
					$sortOrder = 0;
				}
				$sortOrder += 1;
				$optionIds[$productId][$name] = $optionId;
				$sql .= ($first) ? "\n" : ",\n";
				$sql2 .= ($first) ? "\n" : ",\n";
				$first = false;
				$sql .= "($optionId, $productId, $sortOrder )";
				$sql2 .= "($optionId, $productId, $languageId, '" . $this->db->escape($name) . "' )";
			}
		}
		$sql .= ";\n";
		$sql2 .= ";\n";
		if (strlen($sql) > $k + 2) {
			$this->db->query($sql);
			$this->db->query($sql2);
		}
		return true;
	}

	function storeOptionDetailsIntoDatabase(&$options, &$optionIds)
    {
// generate SQL for storing all the option details into the database
		$sql = "INSERT INTO `" . DB_PREFIX . "product_option_value` (`product_option_value_id`, `product_id`, `product_option_id`, `quantity`, `subtract`, `price`, `prefix`, `sort_order`) VALUES ";
		$sql2 = "INSERT INTO `" . DB_PREFIX . "product_option_value_description` (`product_option_value_id`, `product_id`, `language_id`, `name`) VALUES ";
		$k = strlen($sql);
		$first = true;
		foreach ($options as $index => $option) {
			$productOptionValueId = $index + 1;
			$productId = $option['product_id'];
			$optionName = $option['option'];
			$optionId = $optionIds[$productId][$optionName];
			$optionValue = $this->db->escape($option['option_value']);
			$quantity = $option['quantity'];
			$subtract = $option['subtract'];
			$subtract = ((strtolower($subtract) == "true") || (strtoupper($subtract) == "YES") || (strtoupper($subtract) == "ENABLED")) ? 1 : 0;
			$price = $option['price'];
			$prefix = $option['prefix'];
			$sortOrder = $option['sort_order'];
			$sql .= ($first) ? "\n" : ",\n";
			$sql2 .= ($first) ? "\n" : ",\n";
			$first = false;
			$sql .= "($productOptionValueId, $productId, $optionId, $quantity, $subtract, $price, '$prefix', $sortOrder)";
			$sql2 .= "($productOptionValueId, $productId, $languageId, '$optionValue')";
		}
		$sql .= ";\n";
		$sql2 .= ";\n";

// execute the database query
		if (strlen($sql) > $k + 2) {
			$this->db->query($sql);
			$this->db->query($sql2);
		}
		return true;
	}

	function storeOptionsIntoDatabase(&$options) {



// start transaction, remove options
		$sql = "START TRANSACTION;\n";
		$sql .= "DELETE FROM `" . DB_PREFIX . "product_option`;\n";
		$sql .= "DELETE FROM `" . DB_PREFIX . "product_option_description` WHERE language_id='{$this->languageId}';\n";
		$sql .= "DELETE FROM `" . DB_PREFIX . "product_option_value`;\n";
		$sql .= "DELETE FROM `" . DB_PREFIX . "product_option_value_description` WHERE language_id='{$this->languageId}';\n";
		$this->import($sql);

// store option names
		$optionIds = array(); // indexed by product_id and name
		$ok = $this->storeOptionNamesIntoDatabase($options, $optionIds);
		if (!$ok) {
			$this->db->query('ROLLBACK;');
			return false;
		}

// store option details
		$ok = $this->storeOptionDetailsIntoDatabase($options, $optionIds);
		if (!$ok) {
			$this->db->query('ROLLBACK;');
			return false;
		}

		$this->db->query("COMMIT;");
		return true;
	}

	function uploadOptions(&$reader) {
		$data = $reader->getSheet(2);
		$options = array();
		$i = 0;
		$k = $data->getHighestRow();
		$isFirstRow = true;
		for ($i = 0; $i < $k; $i+=1) {
			if ($isFirstRow) {
				$isFirstRow = false;
				continue;
			}
			$productId = trim($this->getCell($data, $i, 1));
			if ($productId == "") {
				continue;
			}
			$languageId = $this->getCell($data, $i, 2);
			$option = $this->getCell($data, $i, 3);
			$optionValue = $this->getCell($data, $i, 4);
			$optionQuantity = $this->getCell($data, $i, 5, '0');
			$optionSubtract = $this->getCell($data, $i, 6, 'false');
			$optionPrice = $this->getCell($data, $i, 7, '0');
			$optionPrefix = $this->getCell($data, $i, 8, '+');
			$sortOrder = $this->getCell($data, $i, 9, '0');
			$options[$i] = array();
			$options[$i]['product_id'] = $productId;
			$options[$i]['language_id'] = $languageId;
			$options[$i]['option'] = $option;
			$options[$i]['option_value'] = $optionValue;
			$options[$i]['quantity'] = $optionQuantity;
			$options[$i]['subtract'] = $optionSubtract;
			$options[$i]['price'] = $optionPrice;
			$options[$i]['prefix'] = $optionPrefix;
			$options[$i]['sort_order'] = $sortOrder;
		}
		return $this->storeOptionsIntoDatabase($options);
	}

	function storeSpecialsIntoDatabase(&$specials) {
		$sql = "START TRANSACTION;\n";
		$sql .= "DELETE FROM `" . DB_PREFIX . "product_special`;\n";
		$this->import($sql);

// find existing customer groups from the database
		$sql = "SELECT * FROM `" . DB_PREFIX . "customer_group`";
		$result = $this->db->query($sql);
		$maxCustomerGroupId = 0;
		$customerGroups = array();
		foreach ($result->rows as $row) {
			$customerGroupId = $row['customer_group_id'];
			$name = $row['name'];
			if (!isset($customerGroups[$name])) {
				$customerGroups[$name] = $customerGroupId;
			}
			if ($maxCustomerGroupId < $customerGroupId) {
				$maxCustomerGroupId = $customerGroupId;
			}
		}

// add additional customer groups into the database
		foreach ($specials as $special) {
			$name = $special['customer_group'];
			if (!isset($customerGroups[$name])) {
				$maxCustomerGroupId += 1;
				$sql = "INSERT INTO `" . DB_PREFIX . "customer_group` (`customer_group_id`, `name`) VALUES ";
				$sql .= "($maxCustomerGroupId, '$name')";
				$sql .= ";\n";
				$this->db->query($sql);
				$customerGroups[$name] = $maxCustomerGroupId;
			}
		}

// store product specials into the database
		$productSpecialId = 0;
		$first = true;
		$sql = "INSERT INTO `" . DB_PREFIX . "product_special` (`product_special_id`,`product_id`,`customer_group_id`,`priority`,`price`,`date_start`,`date_end` ) VALUES ";
		foreach ($specials as $special) {
			$productSpecialId += 1;
			$productId = $special['product_id'];
			$name = $special['customer_group'];
			$customerGroupId = $customerGroups[$name];
			$priority = $special['priority'];
			$price = $special['price'];
			$dateStart = $special['date_start'];
			$dateEnd = $special['date_end'];
			$sql .= ($first) ? "\n" : ",\n";
			$first = false;
			$sql .= "($productSpecialId,$productId,$customerGroupId,$priority,$price,'$dateStart','$dateEnd')";
		}
		if (!$first) {
			$this->db->query($sql);
		}

		$this->db->query("COMMIT;");
		return true;
	}

	function uploadSpecials(&$reader) {
		$data = $reader->getSheet(3);
		$specials = array();
		$i = 0;
		$k = $data->getHighestRow();
		$isFirstRow = true;
		for ($i = 0; $i < $k; $i+=1) {
			if ($isFirstRow) {
				$isFirstRow = false;
				continue;
			}
			$productId = trim($this->getCell($data, $i, 1));
			if ($productId == "") {
				continue;
			}
			$customerGroup = trim($this->getCell($data, $i, 2));
			if ($customerGroup == "") {
				continue;
			}
			$priority = $this->getCell($data, $i, 3, '0');
			$price = $this->getCell($data, $i, 4, '0');
			$dateStart = $this->getCell($data, $i, 5, '0000-00-00');
			$dateEnd = $this->getCell($data, $i, 6, '0000-00-00');
			$specials[$i] = array();
			$specials[$i]['product_id'] = $productId;
			$specials[$i]['customer_group'] = $customerGroup;
			$specials[$i]['priority'] = $priority;
			$specials[$i]['price'] = $price;
			$specials[$i]['date_start'] = $dateStart;
			$specials[$i]['date_end'] = $dateEnd;
		}
		return $this->storeSpecialsIntoDatabase($specials);
	}

	function storeDiscountsIntoDatabase(&$discounts) {
		$sql = "START TRANSACTION;\n";
		$sql .= "DELETE FROM `" . DB_PREFIX . "product_discount`;\n";
		$this->import($sql);

// find existing customer groups from the database
		$sql = "SELECT * FROM `" . DB_PREFIX . "customer_group`";
		$result = $this->db->query($sql);
		$maxCustomerGroupId = 0;
		$customerGroups = array();
		foreach ($result->rows as $row) {
			$customerGroupId = $row['customer_group_id'];
			$name = $row['name'];
			if (!isset($customerGroups[$name])) {
				$customerGroups[$name] = $customerGroupId;
			}
			if ($maxCustomerGroupId < $customerGroupId) {
				$maxCustomerGroupId = $customerGroupId;
			}
		}

// add additional customer groups into the database
		foreach ($discounts as $discount) {
			$name = $discount['customer_group'];
			if (!isset($customerGroups[$name])) {
				$maxCustomerGroupId += 1;
				$sql = "INSERT INTO `" . DB_PREFIX . "customer_group` (`customer_group_id`, `name`) VALUES ";
				$sql .= "($maxCustomerGroupId, '$name')";
				$sql .= ";\n";
				$this->db->query($sql);
				$customerGroups[$name] = $maxCustomerGroupId;
			}
		}

// store product discounts into the database
		$productDiscountId = 0;
		$first = true;
		$sql = "INSERT INTO `" . DB_PREFIX . "product_discount` (`product_discount_id`,`product_id`,`customer_group_id`,`quantity`,`priority`,`price`,`date_start`,`date_end` ) VALUES ";
		foreach ($discounts as $discount) {
			$productDiscountId += 1;
			$productId = $discount['product_id'];
			$name = $discount['customer_group'];
			$customerGroupId = $customerGroups[$name];
			$quantity = $discount['quantity'];
			$priority = $discount['priority'];
			$price = $discount['price'];
			$dateStart = $discount['date_start'];
			$dateEnd = $discount['date_end'];
			$sql .= ($first) ? "\n" : ",\n";
			$first = false;
			$sql .= "($productDiscountId,$productId,$customerGroupId,$quantity,$priority,$price,'$dateStart','$dateEnd')";
		}
		if (!$first) {
			$this->db->query($sql);
		}

		$this->db->query("COMMIT;");
		return true;
	}

	function uploadDiscounts(&$reader) {
		$data = $reader->getSheet(4);
		$discounts = array();
		$i = 0;
		$k = $data->getHighestRow();
		$isFirstRow = true;
		for ($i = 0; $i < $k; $i+=1) {
			if ($isFirstRow) {
				$isFirstRow = false;
				continue;
			}
			$productId = trim($this->getCell($data, $i, 1));
			if ($productId == "") {
				continue;
			}
			$customerGroup = trim($this->getCell($data, $i, 2));
			if ($customerGroup == "") {
				continue;
			}
			$quantity = $this->getCell($data, $i, 3, '0');
			$priority = $this->getCell($data, $i, 4, '0');
			$price = $this->getCell($data, $i, 5, '0');
			$dateStart = $this->getCell($data, $i, 6, '0000-00-00');
			$dateEnd = $this->getCell($data, $i, 7, '0000-00-00');
			$discounts[$i] = array();
			$discounts[$i]['product_id'] = $productId;
			$discounts[$i]['customer_group'] = $customerGroup;
			$discounts[$i]['quantity'] = $quantity;
			$discounts[$i]['priority'] = $priority;
			$discounts[$i]['price'] = $price;
			$discounts[$i]['date_start'] = $dateStart;
			$discounts[$i]['date_end'] = $dateEnd;
		}
		return $this->storeDiscountsIntoDatabase($discounts);
	}

	function storeAdditionalImagesIntoDatabase(&$reader) {
// start transaction
		$sql = "START TRANSACTION;\n";

// delete old additional product images from database
		$sql = "DELETE FROM `" . DB_PREFIX . "product_image`";
		$this->db->query($sql);

// insert new additional product images into database
		$data = & $reader->getSheet(1); // Products worksheet
		$maxImageId = 0;

		$k = $data->getHighestRow();
		for ($i = 1; $i < $k; $i+=1) {
			$productId = trim($this->getCell($data, $i, 1));
			if ($productId == "") {
				continue;
			}
			$imageNames = trim($this->getCell($data, $i, 29));
			$imageNames = trim($this->clean($imageNames, true));
			$imageNames = ($imageNames == "") ? array() : explode(",", $imageNames);
			foreach ($imageNames as $imageName) {
				$imageName = mysql_real_escape_string($imageName);
				$maxImageId += 1;
				$sql = "INSERT INTO `" . DB_PREFIX . "product_image` (`product_image_id`, product_id, `image`) VALUES ";
				$sql .= "($maxImageId,$productId,'$imageName');";
				$this->db->query($sql);
			}
		}

		$this->db->query("COMMIT;");
		return true;
	}

	function uploadImages(&$reader) {
		$ok = $this->storeAdditionalImagesIntoDatabase($reader);
		return $ok;
	}

	/**
	 * Получает значения ячейки таблицы
	 *
	 * @param type $worksheet
	 * @param int $row
	 * @param int $col
	 * @param type $default_val
	 * @return type
	 */
	protected function getCell(&$worksheet, $row, $col, $default_val='') {
		$col -= 1; // we use 1-based, PHPExcel uses 0-based column index
		$row += 1; // we use 0-based, PHPExcel used 1-based row index
		return ($worksheet->cellExistsByColumnAndRow($col, $row)) ? $worksheet->getCellByColumnAndRow($col, $row)->getValue() : $default_val;
	}

	/**
	 * Проверка соответствия таблицы XLS с настройками полей
	 *
	 * @param type $reader
	 * @param type $group
	 * @return boolean
	 */
	function validateHeading(&$reader, $group) {
        return true;

        /*
		$this->loadColumns($group);
		$expected = $this->getColumns(true, true);

		if (empty($expected))
			return false;

		$data = & $reader->getSheet(0);

		$heading = array();
		$k = PHPExcel_Cell::columnIndexFromString($data->getHighestColumn());
		if ($k != count($expected)) {
			return false;
		}
		$i = 0;
		for ($j = 1; $j <= $k; $j+=1) {
			$heading[] = $this->getCell($data, $i, $j);
		}
		$valid = true;
		for ($i = 0; $i < count($expected); $i+=1) {
			if (!isset($heading[$i])) {
				$valid = false;
				break;
			}
			if (strtolower($heading[$i]) != strtolower($expected[$i])) {
				$valid = false;
				break;
			}
		}
		return $valid;
        */
	}

	protected function getStoreIdsForCategories() {
		$sql = "SELECT category_id, store_id FROM `" . DB_PREFIX . "category_to_store` cs;";
		$storeIds = array();
		$result = $this->db->query($sql);
		foreach ($result->rows as $row) {
			$categoryId = $row['category_id'];
			$storeId = $row['store_id'];
			if (!isset($storeIds[$categoryId])) {
				$storeIds[$categoryId] = array();
			}
			if (!in_array($storeId, $storeIds[$categoryId])) {
				$storeIds[$categoryId][] = $storeId;
			}
		}
		return $storeIds;
	}


	function populateCategoriesWorksheet(&$worksheet) {
		// Set the column widths
		$j = 0;
		$i = 0;

        $config = $this->getConfig();
        $categoryFields = isset($config['category']) ? $config['category'] : array();

        foreach($this->columns as $key => $column) {
            if (!isset($categoryFields[$key])) {
                $this->columns[$key]['enabled'] = false;
            }
        }

        $this->load->model('localisation/language');
        $languageList = $this->model_localisation_language->getLanguages();

		$query = "SELECT DISTINCT c.* , ua.keyword";
        foreach ($languageList as $code => $language) {
            $query .= ", description_{$code}.name AS name_{$code}";
            $query .= ", description_{$code}.description AS description_{$code}";
            $query .= ", description_{$code}.meta_description AS meta_description_{$code}";
            $query .= ", description_{$code}.meta_keyword AS meta_keyword_{$code} ";
        }
        $query .= "FROM `" . DB_PREFIX . "category` AS c ";

        foreach ($languageList as $code => $language) {
		    $query .= "INNER JOIN `" . DB_PREFIX . "category_description` AS description_{$code} ON description_{$code}.category_id = c.category_id  AND description_{$code}.language_id='{$language['language_id']}' ";
        }

		$query .= "LEFT JOIN `" . DB_PREFIX . "url_alias` ua ON ua.query=CONCAT('category_id=',c.category_id) ";
		$query .= "ORDER BY c.`parent_id`, `sort_order`, c.`category_id`;";
		$result = $this->db->query($query);

        if (sizeof($result->row) == 0) {
            return false;
        }

        $writeColumns = array();
        foreach ($result->row as $key => $row) {
            foreach($this->columns as $cname => $column ) {
                if (preg_match("/^{$cname}[_a-z]{0,3}$/",$key) && $column['enabled']) {
                    if ( isset($column['multirow']) && $column['multirow'] ) {
                        foreach ($languageList as $code => $language) {
                            $writeColumns[$cname."_{$code}"] = array(
                                'name' => $column['name'],
                                'length' => $column['length'],
                                'format' => $column['format'],
                            );
                        }
                    }
                    else {
                        $writeColumns[$cname] = array(
                            'name' => $column['name'],
                            'length' => $column['length'],
                            'format' => $column['format'],
                        );
                    }
                }
            }
        }

        // write header into XLS file
        foreach ($writeColumns as $colId => $colData) {
            $j++;
            $writeColumns[$colId]['position'] = $j - 1;
            $worksheet->setColumn($j, $j + 1, $colData['length'], $colData['format']);
            $worksheet->writeString($i, $j - 1, $colId, $this->boxFormat);
            $worksheet->writeString($i+1, $j - 1, ($colData['name'] ? $colData['name'] : $colId), $this->boxFormat);
        }
        $i++;
        $worksheet->setRow(0, 1, $this->boxFormat); // set first row height 1 px (ID fields)
        $worksheet->setRow($i, 30, $this->boxFormat);
        $i++; $j = 0;
        $storeIds = $this->getStoreIdsForCategories();

        foreach ($result->rows as $row) {
            foreach ($writeColumns as $colName => $col) {
                $worksheet->write($i, $col['position'], $row[$colName]);
            }
            $i++;
        }
	}

	function getStoreIdsForProducts() {
		$sql = "SELECT product_id, store_id FROM `" . DB_PREFIX . "product_to_store` ps;";
		$storeIds = array();
		$result = $this->db->query($sql);
		foreach ($result->rows as $row) {
			$productId = $row['product_id'];
			$storeId = $row['store_id'];
			if (!isset($storeIds[$productId])) {
				$storeIds[$productId] = array();
			}
			if (!in_array($storeId, $storeIds[$productId])) {
				$storeIds[$productId][] = $storeId;
			}
		}
		return $storeIds;
	}

	function populateProductsWorksheet(&$worksheet) {

        $this->load->model('localisation/language');
        $languageList = $this->model_localisation_language->getLanguages();

		// Set the column widths
        $config = $this->getConfig();
        $productFields = $config['product'];

        $i = 1;
        foreach ($this->columns as $key => $column ) {
            if ( !isset($productFields[$key]) ) {
                $this->columns[$key]['enabled'] = false;
            }
            else {
                $this->columns[$key]['enabled'] = true;
                $i++;
            }
        }

		$j = 1;
		$i = 0;
        $colCount = 0;
        $writeColumns = array();
		foreach ($this->columns as $colId => $colData) {
			if ($colData['enabled'] ) {
                if (!$colData['multirow']) {
			        $worksheet->setColumn($j, $j + 1, $colData['length'], $colData['format']);
                    $worksheet->writeString($i, $j - 1, $colId, $this->boxFormat);
			        $worksheet->writeString($i+1, $j - 1, ($colData['name'] ? $colData['name'] : $colId), $this->boxFormat);
                    $writeColumns[$colId] = $colData;
                    $writeColumns[$colId]['position'] = $j - 1;
                    $colCount++;
                    $j++;
                }
                else {
                    foreach ($languageList as $code => $language ) {
                        $worksheet->setColumn($j, $j + 1, $colData['length'], $colData['format']);
                        $worksheet->writeString($i, $j - 1, "{$colId}_{$code}", $this->boxFormat);
                        $worksheet->writeString($i+1, $j - 1, "{$colData['name']}({$code})", $this->boxFormat);
                        $writeColumns["{$colId}_{$code}"] = $colData;
                        $writeColumns["{$colId}_{$code}"]['position'] = $j - 1;
                        $colCount++;
                        $j++;
                    }
                }
            }
		}

		$worksheet->setRow($i, 1, $this->boxFormat);
        $worksheet->setRow($i+1, 30, $this->boxFormat);

        $query = "SELECT _p.product_id AS product_id,";
        $query .= "GROUP_CONCAT( _ptc.category_id ) AS categories, ";
        $query .= "_p.sku AS sku, ";
        $query .= "_p.location AS location, ";
        $query .= "_p.quantity AS quantity, ";
        $query .= "_p.model AS model, ";
        $query .= "_m.name AS manufacturer, ";
        $query .= "GROUP_CONCAT(_pi.image) AS image, ";
        $query .= "_p.shipping AS shipping, ";
        $query .= "_p.price AS price, ";
        $query .= "_p.weight AS weight, ";
        $query .= "_wcd.unit AS unit, ";
        $query .= "_p.length AS length, ";
        $query .= "_p.width AS width, ";
        $query .= "_p.height AS height, ";
        $query .= "_lcd.unit AS length_unit, ";
        $query .= "_p.status AS status, ";
        $query .= "_p.tax_class_id AS tax_class_id, ";
        $query .= "_p.viewed AS viewed, ";
        $query .= "_ua.keyword AS keyword, ";
        $query .= "_p.stock_status_id AS stock_status_id, ";

        foreach ($languageList as $code => $language) {
            $query .= "_pd_{$code}.description AS description_{$code}, ";
            $query .= "_pd_{$code}.name AS name_{$code}, ";
            $query .= "_pd_{$code}.meta_description AS meta_description_{$code}, ";
            $query .= "_pd_{$code}.meta_keyword AS meta_keyword_{$code}, ";
            $query .= "_pd_{$code}.tag AS tags_{$code}, ";
        }

        $query .= "GROUP_CONCAT(_pts.store_id SEPARATOR ',') AS store_ids, ";
        $query .= "GROUP_CONCAT(_pr.related_id SEPARATOR ',') AS related, ";
        $query .= "_p.sort_order AS sort_order, ";
        $query .= "_p.subtract AS subtract, ";
        $query .= "_p.minimum AS minimum, ";
        $query .= "_p.date_added AS date_added, ";
        $query .= "_p.date_modified AS date_modified, ";
        $query .= "_p.date_available AS date_available ";
        $query .= "FROM `" . DB_PREFIX . "product` AS _p ";
        $query .= "LEFT JOIN `" . DB_PREFIX . "product_to_category` AS _ptc ON _ptc.product_id = _p.product_id ";
        $query .= "LEFT JOIN `" . DB_PREFIX . "product_to_store` AS _pts ON _pts.product_id = _p.product_id ";
        $query .= "LEFT JOIN `" . DB_PREFIX . "product_image` AS _pi ON _pi.product_id = _p.product_id ";
        $query .= "LEFT JOIN `" . DB_PREFIX . "weight_class_description` AS _wcd  ON _wcd.weight_class_id = _p.weight_class_id AND _wcd.language_id = '{$this->languageId}' ";
        $query .= "LEFT JOIN `" . DB_PREFIX . "length_class_description` AS _lcd  ON _lcd.length_class_id = _p.length_class_id AND _wcd.language_id = '{$this->languageId}' ";
        $query .= "LEFT JOIN `" . DB_PREFIX . "url_alias` AS _ua ON _ua.query=CONCAT('product_id=',_p.product_id) ";
        $query .= "LEFT JOIN `" . DB_PREFIX . "manufacturer` AS _m ON _m.manufacturer_id = _p.manufacturer_id ";
        $query .= "LEFT JOIN `" . DB_PREFIX . "product_related` _pr ON _pr.product_id=_p.product_id ";
        foreach ($languageList as $code => $language) {
            $query .= "LEFT JOIN `" . DB_PREFIX . "product_description` _pd_{$code} ON _pd_{$code}.product_id=_p.product_id AND _pd_{$code}.language_id = '{$language['language_id']}' ";
        };
        $query .= "GROUP BY _p.product_id";

        $result = $this->db->query($query);

        $i++;

        foreach ($result->rows as $row) {
            $i++;
            if (isset($row['categories']))
                $row['categories'] = implode(array_flip(array_flip(explode(',', $row['categories']))), ',');
            if (isset($row['image']))
                $row['image'] = implode(array_flip(array_flip(explode(',', $row['image']))), ',');
            if (isset($row['store_ids']))
                $row['store_ids'] = implode(array_flip(array_flip(explode(',', $row['store_ids']))), ',');
            if (isset($row['related']))
                $row['related'] = implode(array_flip(array_flip(explode(',', $row['related']))), ',');

            foreach ($row as $cellName => $cell) {
                if (isset($writeColumns[$cellName]) ) {
                    $worksheet->write($i, $writeColumns[$cellName]['position'], $cell, $this->textFormat);
                }
            }
        }
	}

	function populateOptionsWorksheet(&$worksheet) {
// Set the column widths
		$j = 0;
		$worksheet->setColumn($j, $j++, max(strlen('product_id'), 4) + 1);
		$worksheet->setColumn($j, $j++, max(strlen('language_id'), 2) + 1);
		$worksheet->setColumn($j, $j++, max(strlen('option'), 30) + 1);
		$worksheet->setColumn($j, $j++, max(strlen('option_value'), 30) + 1);
		$worksheet->setColumn($j, $j++, max(strlen('quantity'), 4) + 1);
		$worksheet->setColumn($j, $j++, max(strlen('subtract'), 5) + 1, $this->textFormat);
		$worksheet->setColumn($j, $j++, max(strlen('price'), 10) + 1, $this->priceFormat);
		$worksheet->setColumn($j, $j++, max(strlen('prefix'), 5) + 1, $this->textFormat);
		$worksheet->setColumn($j, $j++, max(strlen('sort_order'), 5) + 1);

// The options headings row
		$i = 0;
		$j = 0;
		$worksheet->writeString($i, $j++, 'product_id', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'language_id', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'option', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'option_value', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'quantity', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'subtract', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'price', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'prefix', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'sort_order', $this->boxFormat);
		$worksheet->setRow($i, 30, $this->boxFormat);

// The actual options data
		$i++;
		$j = 0;
		$query = "SELECT DISTINCT p.product_id, ";
		$query .= "  pod.name AS option_name, ";
		$query .= "  po.sort_order AS option_sort_order, ";
		$query .= "  povd.name AS option_value, ";
		$query .= "  pov.quantity AS option_quantity, ";
		$query .= "  pov.subtract AS option_subtract, ";
		$query .= "  pov.price AS option_price, ";
		$query .= "  pov.prefix AS option_prefix, ";
		$query .= "  pov.sort_order AS sort_order ";
		$query .= "FROM `" . DB_PREFIX . "product` p ";
		$query .= "INNER JOIN `" . DB_PREFIX . "product_description` pd ON p.product_id=pd.product_id ";
		$query .= "  AND pd.language_id='{$this->languageId}' ";
		$query .= "INNER JOIN `" . DB_PREFIX . "product_option` po ON po.product_id=p.product_id ";
		$query .= "INNER JOIN `" . DB_PREFIX . "product_option_description` pod ON pod.product_option_id=po.product_option_id ";
		$query .= "  AND pod.product_id=po.product_id ";
		$query .= "  AND pod.language_id='{$this->languageId}' ";
		$query .= "INNER JOIN `" . DB_PREFIX . "product_option_value` pov ON pov.product_option_id=po.product_option_id ";
		$query .= "INNER JOIN `" . DB_PREFIX . "product_option_value_description` povd ON povd.product_option_value_id=pov.product_option_value_id ";
		$query .= "  AND povd.language_id='{$this->languageId}' ";
		$query .= "ORDER BY product_id, option_sort_order, sort_order;";
		$result = $this->db->query($query);
		foreach ($result->rows as $row) {
			$worksheet->write($i, $j++, $row['product_id']);
			$worksheet->write($i, $j++, $languageId);
			$worksheet->writeString($i, $j++, $row['option_name']);
			$worksheet->writeString($i, $j++, $row['option_value']);
			$worksheet->write($i, $j++, $row['option_quantity']);
			$worksheet->write($i, $j++, ($row['option_subtract'] == 0) ? "false" : "true", $this->textFormat);
			$worksheet->write($i, $j++, $row['option_price'], $this->priceFormat);
			$worksheet->writeString($i, $j++, $row['option_prefix'], $this->textFormat);
			$worksheet->write($i, $j++, $row['sort_order']);
			$i++;
			$j = 0;
		}
	}

	function populateSpecialsWorksheet(&$worksheet) {
// Set the column widths
		$j = 0;
		$worksheet->setColumn($j, $j++, strlen('product_id') + 1);
		$worksheet->setColumn($j, $j++, strlen('customer_group') + 1);
		$worksheet->setColumn($j, $j++, strlen('priority') + 1);
		$worksheet->setColumn($j, $j++, max(strlen('price'), 10) + 1, $this->priceFormat);
		$worksheet->setColumn($j, $j++, max(strlen('date_start'), 19) + 1, $this->textFormat);
		$worksheet->setColumn($j, $j++, max(strlen('date_end'), 19) + 1, $this->textFormat);

// The heading row
		$i = 0;
		$j = 0;
		$worksheet->writeString($i, $j++, 'product_id', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'customer_group', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'priority', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'price', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'date_start', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'date_end', $this->boxFormat);
		$worksheet->setRow($i, 30, $this->boxFormat);

// The actual product specials data
		$i++;
		$j = 0;
		$query = "SELECT ps.*, cg.name FROM `" . DB_PREFIX . "product_special` ps ";
		$query .= "LEFT JOIN `" . DB_PREFIX . "customer_group` cg ON cg.customer_group_id=ps.customer_group_id ";
		$query .= "ORDER BY ps.product_id, cg.name";
		$result = $this->db->query($query);
		foreach ($result->rows as $row) {
			$worksheet->write($i, $j++, $row['product_id']);
			$worksheet->write($i, $j++, $row['name']);
			$worksheet->write($i, $j++, $row['priority']);
			$worksheet->write($i, $j++, $row['price'], $this->priceFormat);
			$worksheet->write($i, $j++, $row['date_start'], $this->textFormat);
			$worksheet->write($i, $j++, $row['date_end'], $this->textFormat);
			$i++;
			$j = 0;
		}
	}

	function populateDiscountsWorksheet(&$worksheet) {
// Set the column widths
		$j = 0;
		$worksheet->setColumn($j, $j++, strlen('product_id') + 1);
		$worksheet->setColumn($j, $j++, strlen('customer_group') + 1);
		$worksheet->setColumn($j, $j++, strlen('quantity') + 1);
		$worksheet->setColumn($j, $j++, strlen('priority') + 1);
		$worksheet->setColumn($j, $j++, max(strlen('price'), 10) + 1, $this->priceFormat);
		$worksheet->setColumn($j, $j++, max(strlen('date_start'), 19) + 1, $this->textFormat);
		$worksheet->setColumn($j, $j++, max(strlen('date_end'), 19) + 1, $this->textFormat);

// The heading row
		$i = 0;
		$j = 0;
		$worksheet->writeString($i, $j++, 'product_id', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'customer_group', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'quantity', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'priority', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'price', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'date_start', $this->boxFormat);
		$worksheet->writeString($i, $j++, 'date_end', $this->boxFormat);
		$worksheet->setRow($i, 30, $this->boxFormat);

// The actual product discounts data
		$i++;
		$j = 0;
		$query = "SELECT pd.*, cg.name FROM `" . DB_PREFIX . "product_discount` pd ";
		$query .= "LEFT JOIN `" . DB_PREFIX . "customer_group` cg ON cg.customer_group_id=pd.customer_group_id ";
		$query .= "ORDER BY pd.product_id, cg.name";
		$result = $this->db->query($query);
		foreach ($result->rows as $row) {
			$worksheet->write($i, $j++, $row['product_id']);
			$worksheet->write($i, $j++, $row['name']);
			$worksheet->write($i, $j++, $row['quantity']);
			$worksheet->write($i, $j++, $row['priority']);
			$worksheet->write($i, $j++, $row['price'], $this->priceFormat);
			$worksheet->write($i, $j++, $row['date_start'], $this->textFormat);
			$worksheet->write($i, $j++, $row['date_end'], $this->textFormat);
			$i++;
			$j = 0;
		}
	}

	/**
	 * Load default language
	 *
	 * @return type
	 */
	protected function loadDefaultLanguage() {
		$code = $this->config->get('config_language');
		$sql = "SELECT language_id FROM `" . DB_PREFIX . "language` WHERE code = '$code'";
		$result = $this->db->query($sql);
		$this->languageId = ($result->row && $result->row['language_id']) ? $result->row['language_id'] : 1;
	}

	/**
	 * Log writer
	 * @param type $text
	 */
	protected function log($text) {
		if ($text = trim($text)) {
			error_log(date('Y-m-d H:i:s - ', time()) . "Export/Import: {$text}\n", 3, DIR_LOGS . "error.txt");
		}
	}

}

?>
