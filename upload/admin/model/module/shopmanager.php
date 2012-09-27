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
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true,
                'readonly' => true,
            ),
            "parent_id" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            "name" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            "sort_order" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            "image" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            "date_added" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            "date_modified" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            "language_id" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            "keyword" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            "description" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            "meta_description" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            "meta_keyword" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            "store_ids" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            "status" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            )
        );


        /*return array(
            'category.category_id' =>  array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true,
            ),
            'category.parent_id' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            'category.sort_order' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true,
            ),
            'category_description.name' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true,
            ),
            'category.image' => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true,
            ),
            'category.date_added' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true,
            ),
            'category.date_modified' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            'category_description.language_id' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            'category_description.language_id' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            'url_alias.keyword' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            'category_description.description' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            'category_description.meta_description' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            'category_description.meta_keyword' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            'category_to_store.store_ids' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            'category.status' => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
        );*/
    }

    /**
     *  product fields list
     *
     */
    public function getProductFildsList()
    {
        return array(
            "product_id" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            "name" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            "categories" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            "sku" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            "location" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            "quantity" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            "model" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            "manufacturer" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            "image" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            "shipping" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            "price" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->priceFormat,
                'enabled' => true
            ),
            "weight" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->weightFormat,
                'enabled' => false
            ),
            "unit" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            "length" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            "width" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            "height" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            "length_unit" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            "status" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            "tax_class_id" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            "viewed" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            "language_id" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            "keyword" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            "description" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            "meta_description" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            "meta_keyword" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            "images" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            "stock_status_id" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => true
            ),
            "store_ids" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            "related" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            "tags" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            "sort_order" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            "subtract" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => false
            ),
            "minimum" => array(
                'name' => null,
                'length' => 20,
                'format' => null,
                'enabled' => false
            ),
            "date_added" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            "date_modified" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
            ),
            "date_available" => array(
                'name' => null,
                'length' => 20,
                'format' => $this->textFormat,
                'enabled' => true
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
	public function export($group) {

		$this->init();
		ini_set("memory_limit", "128M");
		require_once 'Spreadsheet/Excel/Writer.php';

        // Creating a workbook
		$workbook = new Spreadsheet_Excel_Writer();

//		$workbook->setTempDir(DIR_CACHE);

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

		$worksheet->freezePanes(array(1, 1, 1, 1));

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

	protected function updateManufacturersInDB(&$products, &$manufacturerIds) {
		// find all manufacturers already stored in the database
		$sql = "SELECT `manufacturer_id`, `name` FROM `" . DB_PREFIX . "manufacturer`;";
		$result = $this->db->query($sql);

		if ($result->rows) {
			foreach ($result->rows as $row) {
				$manufacturerId = $row['manufacturer_id'];
				$name = $row['name'];
				if (!isset($manufacturerIds[$name])) {
					$manufacturerIds[$name] = $manufacturerId;
				} else if ($manufacturerIds[$name] < $manufacturerId) {
					$manufacturerIds[$name] = $manufacturerId;
				}
			}
		}

		// add newly introduced manufacturers to the database
		$maxManufacturerId = 0;
		foreach ($manufacturerIds as $manufacturerId) {
			$maxManufacturerId = max($maxManufacturerId, $manufacturerId);
		}
		$sql = "INSERT INTO `" . DB_PREFIX . "manufacturer` (`manufacturer_id`, `name`, `image`, `sort_order`) VALUES ";
		$k = strlen($sql);
		$first = true;
		foreach ($products as $product) {
			$manufacturerName = $product['manufacturer'];
			if (empty($manufacturerName))
				continue;

			if (!isset($manufacturerIds[$manufacturerName])) {
				$maxManufacturerId += 1;
				$manufacturerId = $maxManufacturerId;
				$manufacturerIds[$manufacturerName] = $manufacturerId;
				$sql .= ($first) ? "\n" : ",\n";
				$first = false;
				$sql .= "($manufacturerId, '" . $this->db->escape($manufacturerName) . "', '', 0)";
			}
		}
		$sql .= ";\n";
		if (strlen($sql) > $k + 2) {
			$this->db->query($sql);
		}

		// populate manufacturer_to_store table
		if ($this->isEnabled('store_ids')) {
			$storeIdsForManufacturers = array();
			foreach ($products as $product) {
				$manufacturerName = $product['manufacturer'];
				if (empty($manufacturerName))
					continue;
				$manufacturerId = $manufacturerIds[$manufacturerName];
				$storeIds = $product['store_ids'];
				if (!isset($storeIdsForManufacturers[$manufacturerId])) {
					$storeIdsForManufacturers[$manufacturerId] = array();
				}
				foreach ($storeIds as $storeId) {
					if ($storeId == '')
						continue;
					if (!in_array($storeId, $storeIdsForManufacturers[$manufacturerId])) {
						$storeIdsForManufacturers[$manufacturerId][] = $storeId;
						$sql = "INSERT INTO `" . DB_PREFIX . "manufacturer_to_store` (`manufacturer_id`,`store_id`) ";
						$sql .= "VALUES ({$manufacturerId},{$storeId}) ";
						$sql .= "ON DUPLICATE KEY UPDATE `manufacturer_id`='{$manufacturerId}',`store_id`='{$storeId}';";
						$this->db->query($sql);
					}
				}
			}
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

	protected function updateProductsInDB($products) {
		$this->import("START TRANSACTION;\n");

		// store or update manufacturers
		$manufacturerIds = array();
		$ok = $this->updateManufacturersInDB($products, $manufacturerIds);

		if (!$ok) {
			$this->db->query('ROLLBACK;');
			return false;
        }



		// get weight classes
		$weightClassIds = $this->getWeightClassIds();

		// get length classes
		$lengthClassIds = $this->getLengthClassIds();

		// generate and execute SQL for storing the products
		foreach ($products as $_key => $product) {
			if ($this->isEnabled('product_id') && $product['product_id']) {
				$productId = $product['product_id'];
			} else {
				continue;
			}

			$sql = "UPDATE `" . DB_PREFIX . "product` SET ";

            $flag = false;

			if ($this->isEnabled('quantity')) {
				$sql .= "`quantity`='{$product['quantity']}',";
                $flag = true;
			}
			if ($this->isEnabled('sku')) {
				$sql .= "`sku`='{$this->db->escape($product['sku'])}', ";
                $flag = true;
			}
			if ($this->isEnabled('location')) {
				$sql .= "`location`='{$this->db->escape($product['location'])}', ";
                $flag = true;
			}
			if ($this->isEnabled('stock_status_id')) {
				$sql .= "`stock_status_id`='" . (int) $product['stock_status_id'] . "',";
                $flag = true;
			}
			if ($this->isEnabled('model')) {
				$sql .= "`model`='{$this->db->escape($product['model'])}', ";
                $flag = true;
			}
			if ($this->isEnabled('manufacturer')) {
				$manufacturerName = $product['manufacturer'];
                $mId = (int) isset($manufacturerIds[$manufacturerName]) ? $manufacturerIds[$manufacturerName] : 'NULL';
				$sql .= "`manufacturer_id`=" . $mId  . ",";
                $flag = true;
			}
			if ($this->isEnabled('image')) {
				$sql .= "`image`='{$product['image']}',";
                $flag = true;
			}
			if ($this->isEnabled('shipping')) {
				$shipping = $product['shipping'];
				$shipping = ((strtoupper($shipping) == "YES") || (strtoupper($shipping) == "Y")) ? 1 : 0;
				$sql .= "`shipping`='{$shipping}',";
                $flag = true;
			}
			if ($this->isEnabled('price')) {
				$sql .= "`price`='" . trim($product['price']) . "',";
                $flag = true;
			}
			if ($this->isEnabled('date_added')) {
				$product['date_added'] = ($product['date_added'] == 'NOW()') ? "{$product['date_added']}" : "'{$product['date_added']}'";
				$sql .= "`date_added`={$product['date_added']}, ";
                $flag = true;
			}
			if ($this->isEnabled('date_modified')) {
				$product['date_modified'] = ($product['date_modified'] == 'NOW()') ? "{$product['date_modified']}" : "'{$product['date_modified']}'";
				$sql .= "`date_modified`={$product['date_modified']}, ";
                $flag = true;
			}
			if ($this->isEnabled('date_available')) {
				$product['date_available'] = ($product['date_available'] == 'NOW()') ? "{$product['date_available']}" : "'{$product['date_available']}'";
				$sql .= "`date_available`={$product['date_available']}, ";
                $flag = true;
			}
			if ($this->isEnabled('weight')) {
				$weight = empty($product['weight']) ? 0 : $product['weight'];
				$sql .= "`weight`='{$weight}', ";
                $flag = true;
			}
			if ($this->isEnabled('unit')) {
				$unit = $product['unit'];
				$weightClassId = (isset($weightClassIds[$unit])) ? $weightClassIds[$unit] : 0;
				$sql .= "`weight_class_id`='{$weightClassId}', ";
                $flag = true;
			}
			if ($this->isEnabled('status')) {
				$status = $product['status'];
				$status = ((strtolower($status) == "true") || (strtoupper($status) == "YES") || (strtoupper($status) == "ENABLED")) ? 1 : 0;
				$sql .= "`status`='{$status}', ";
                $flag = true;
			}
			if ($this->isEnabled('stock_status_id')) {
				$manufacturerName = $product['manufacturer'];
				$sql .= "`manufacturer_id`='" . (int) $manufacturerIds[$manufacturerName] . "',";
                $flag = true;
			}
			if ($this->isEnabled('tax_class_id')) {
				$sql .= "`tax_class_id`='" . (int) $product['tax_class_id'] . "',";
                $flag = true;
			}
			if ($this->isEnabled('viewed')) {
                $flag = true;
				$sql .= "`viewed`='" . (int) $product['viewed'] . "',";
			}
			if ($this->isEnabled('length')) {
				$sql .= "`length`='" . trim($product['length']) . "',";
                $flag = true;
			}
			if ($this->isEnabled('width')) {
				$sql .= "`width`='" . trim($product['width']) . "',";
                $flag = true;
			}
			if ($this->isEnabled('height')) {
				$sql .= "`height`='" . trim($product['height']) . "',";
                $flag = true;
			}
			if ($this->isEnabled('length_unit')) {
				$lengthUnit = $product['length_unit'];
				$lengthClassId = (isset($lengthClassIds[$lengthUnit])) ? $lengthClassIds[$lengthUnit] : 0;
				$sql .= "`length_class_id`='{$lengthClassId}', ";
                $flag = true;
			}
			if ($this->isEnabled('sort_order')) {
				$sql .= "`sort_order`='" . (int) $product['sort_order'] . "',";
                $flag = true;
			}
			if ($this->isEnabled('subtract')) {
				$subtract = $product['subtract'];
				$subtract = ((strtolower($subtract) == "true") || (strtoupper($subtract) == "YES") || (strtoupper($subtract) == "ENABLED")) ? 1 : 0;
				$sql .= "`subtract`='" . trim($subtract) . "',";
                $flag = true;
			}
			if ($this->isEnabled('minimum')) {
				$sql .= "`minimum`='" . trim($product['minimum']) . "',";
                $flag = true;
			}
			$sql = rtrim(trim($sql), ',') . " WHERE `product_id` = {$productId};";

            if ($flag) {
			    $this->db->query($sql);
            }

			if ($this->isEnabled('language_id') && $product['language_id']) {
				$languageId = $product['language_id'];
			} else {
                $languageId = $this->languageId;
				//continue;
			}

            $flag = false;
			$sql = "UPDATE `" . DB_PREFIX . "product_description` SET ";
			if ($this->isEnabled('name')) {
				$sql .= "`name`='" . $this->db->escape($product['name']) . "',";
                $flag = true;
			}
			if ($this->isEnabled('description')) {
				$sql .= "`description`='" . $this->db->escape($product['description']) . "',";
                $flag = true;
			}
			if ($this->isEnabled('meta_description')) {
				$sql .= "`meta_description`='" . $this->db->escape($product['meta_description']) . "',";
                $flag = true;
			}
			if ($this->isEnabled('meta_keyword')) {
				$sql .= "`meta_keyword`='" . $this->db->escape($product['meta_keyword']) . "',";
                $flag = true;
			}

            if ($this->isEnabled('tags')) {
                $sql .= "`tag`='" . $this->db->escape($product['tags']) . "',";
                $flag = true;
            }

			$sql = rtrim(trim($sql), ',');
			$sql .= " WHERE `product_id` = {$productId}";
			$sql .= " AND `language_id` = {$languageId};";

            if ($flag) {
			    $this->db->query($sql);
            }

			if ($this->isEnabled('categories') && count($product['categories']) > 0) {
                $this->db->query("DELETE FROM `" . DB_PREFIX . "product_to_category` WHERE `product_id`='{$productId}';");
				foreach ($product['categories'] as $categoryId) {
					if (!($categoryId = (int) $categoryId))
						continue;
					$sql = "INSERT INTO `" . DB_PREFIX . "product_to_category` (`product_id`,`category_id`) ";
					$sql .= "VALUES ('{$productId}','{$categoryId}')";

				    $this->db->query($sql);
				}
			}

			if ($this->isEnabled('keyword')) {
				$sql = "INSERT INTO `" . DB_PREFIX . "url_alias` (`query`,`keyword`) ";
				$sql .= "VALUES ('product_id={$productId}','{$product['keyword']}') ";
				$sql .= "ON DUPLICATE KEY UPDATE `query`='product_id={$productId}',`keyword`='{$product['keyword']}';";
				$this->db->query($sql);
			}
			if ($this->isEnabled('store_ids')) {
				foreach ($product['store_ids'] as $storeId) {
					if ($storeId == '')
						continue;
					$sql = "INSERT INTO `" . DB_PREFIX . "product_to_store` (`product_id`,`store_id`) ";
					$sql .= "VALUES ({$productId},{$storeId}) ";
					$sql .= "ON DUPLICATE KEY UPDATE `product_id`='{$productId}',`store_id`='{$storeId}';";
					$this->db->query($sql);
				}
			}

			continue;

			$keyword = $this->db->escape($product[29]);

			$tags = array();
			foreach ($product[35] as $tag) {
				$tags[] = $this->db->escape($tag);
			}
			if (count($related) > 0) {
				$sql = "INSERT INTO `" . DB_PREFIX . "product_related` (`product_id`,`related_id`) VALUES ";
				$first = true;
				foreach ($related as $relatedId) {
					if (!($relatedId = (int) $relatedId))
						continue;
					$sql .= ($first) ? "\n" : ",\n";
					$first = false;
					$sql .= "($productId,$relatedId)";
				}
				$sql .= ";";
				$this->db->query($sql);
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
        $this->columns = $this->getProductFildsList();
        $this->buildFieldRelations($reader, 'product');
        $this->columns = $this->relation;

		$defaultWeightUnit = $this->getDefaultWeightUnit();
		$defaultMeasurementUnit = $this->getDefaultMeasurementUnit();
		$defaultStockStatusId = $this->config->get('config_stock_status_id');

		$data = $reader->getSheet(0);
		$products = array();
		$isFirstRow = true;
		$i = 0;
		$j = 1;
		$k = $data->getHighestRow();
		for ($i = 0; $i < $k; $i+=1) {
			if ($isFirstRow) {
				$isFirstRow = false;
				continue;
			}
			$product = array();
			$productId = trim($this->getCell($data, $i, $j++));

			if ($this->isEnabled('product_id') && $productId) {
				$product['product_id'] = $productId;
			} else {
				continue;
			}
			if ($this->isEnabled('name')) {
				$name = trim($this->getCell($data, $i, $this->columns['name']['position']));
				$product['name'] = htmlentities($name, ENT_QUOTES, $this->detect_encoding($name));
			}
			if ($this->isEnabled('categories')) {
				$categories = $this->getCell($data, $i, $this->columns['categories']['position']);
				$categories = trim($this->clean($categories, false));
				$product['categories'] = ($categories == "") ? array() : explode(",", $categories);
				if ($product['categories'] === false) {
					$product['categories'] = array();
				}
			}
			if ($this->isEnabled('sku')) {
				$product['sku'] = trim($this->getCell($data, $i, $this->columns['sku']['position'], '0'));
			}
			if ($this->isEnabled('location')) {
				$product['location'] = trim($this->getCell($data, $i, $this->columns['location']['position'], '0'));
			}
			if ($this->isEnabled('quantity')) {
				$product['quantity'] = trim($this->getCell($data, $i, $this->columns['quantity']['position'], '0'));
			}
			if ($this->isEnabled('model')) {
				$product['model'] = $this->getCell($data, $i, $this->columns['model']['position'], '   ');
			}
			if ($this->isEnabled('manufacturer')) {
				$product['manufacturer'] = trim($this->getCell($data, $i, $this->columns['manufacturer']['position']));
			}
			if ($this->isEnabled('image')) {
				$product['image'] = mysql_real_escape_string($this->getCell($data, $i, $this->columns['image']['position']));
			}
			if ($this->isEnabled('shipping')) {
				$product['shipping'] = trim($this->getCell($data, $i, $this->columns['shipping']['position'], 'yes'));
			}
			if ($this->isEnabled('price')) {
				$product['price'] = trim($this->getCell($data, $i, $this->columns['price']['position'], '0.00'));
			}
			if ($this->isEnabled('weight')) {
				$product['weight'] = trim($this->getCell($data, $i, $this->columns['weight']['position'], '0'));
			}
			if ($this->isEnabled('unit')) {
				$product['unit'] = trim($this->getCell($data, $i, $this->columns['unit']['position'], $defaultWeightUnit));
			}
			if ($this->isEnabled('length')) {
				$product['length'] = trim($this->getCell($data, $i, $this->columns['length']['position'], '0'));
			}
			if ($this->isEnabled('width')) {
				$product['width'] = trim($this->getCell($data, $i, $this->columns['width']['position'], '0'));
			}
			if ($this->isEnabled('height')) {
				$product['height'] = trim($this->getCell($data, $i, $this->columns['height']['position'], '0'));
			}
			if ($this->isEnabled('length_unit')) {
				$product['length_unit'] = trim($this->getCell($data, $i, $this->columns['length_unit']['position'], $defaultMeasurementUnit));
			}
			if ($this->isEnabled('status')) {
				$product['status'] = trim($this->getCell($data, $i, $this->columns['status']['position'], 'true'));
			}
			if ($this->isEnabled('tax_class_id')) {
				$product['tax_class_id'] = trim($this->getCell($data, $i, $this->columns['tax_class_id']['position'], '0'));
			}
			if ($this->isEnabled('viewed')) {
				$product['viewed'] = trim($this->getCell($data, $i, $this->columns['viewed']['position'], '0'));
			}
            /**
             * @todo Костыль детектед. Язык захардкожен, чтобы не выводить его в таблицу XLS
             */
			$langId = $this->languageId;//trim($this->getCell($data, $i, $j++, $this->languageId));
			if ($this->isEnabled('language_id') && $langId && ($langId == $this->languageId)) {
				$product['language_id'] = $langId;
			} else {
                $product['language_id'] = $this->languageId;
				//continue;
			}
			if ($this->isEnabled('keyword')) {
				$product['keyword'] = trim($this->getCell($data, $i, $this->columns['keyword']['position']));
			}
			if ($this->isEnabled('description')) {
				$description = trim($this->getCell($data, $i, $this->columns['description']['position']));
				$product['description'] = htmlentities($description, ENT_QUOTES, $this->detect_encoding($description));
			}
			if ($this->isEnabled('meta_description')) {
				$meta_description = trim($this->getCell($data, $i, $this->columns['meta_description']['position']));
				$product['meta_description'] = htmlentities($meta_description, ENT_QUOTES, $this->detect_encoding($meta_description));
			}
			if ($this->isEnabled('meta_keyword')) {
				$meta_keyword = trim($this->getCell($data, $i, $this->columns['meta_keyword']['position']));
				$product['meta_keyword'] = htmlentities($meta_keyword, ENT_QUOTES, $this->detect_encoding($meta_keyword));
			}
			if ($this->isEnabled('images')) {
				$product['images'] = trim($this->getCell($data, $i, $this->columns['images']['position']));
			}
			if ($this->isEnabled('stock_status_id')) {
				$product['stock_status_id'] = trim($this->getCell($data, $i, $this->columns['stock_status_id']['position'], $defaultStockStatusId));
			}
			if ($this->isEnabled('store_ids')) {
				$storeIds = $this->getCell($data, $i, $this->columns['store_ids']['position']);
				$storeIds = trim($this->clean($storeIds, false));
				$product['store_ids'] = ($storeIds == "") ? array() : explode(",", $storeIds);
				if ($product['store_ids'] === false) {
					$product['store_ids'] = array();
				}
			}
			if ($this->isEnabled('related')) {
				$related = trim($this->getCell($data, $i, $this->columns['related']['position']));
				$product['related'] = ($related == "") ? array() : explode(",", $related);
				if ($product['related'] === false) {
					$product['related'] = array();
				}
			}
			if ($this->isEnabled('tags')) {
				$tags = trim($this->getCell($data, $i, $this->columns['tags']['position']));
				$product['tags'] = ($tags == "") ? array() : explode(",", $tags);
				if ($product['tags'] === false) {
					$product['tags'] = array();
				}
			}
			if ($this->isEnabled('sort_order')) {
				$product['sort_order'] = trim($this->getCell($data, $i, $this->columns['sort_order']['position'], '0'));
			}
			if ($this->isEnabled('subtract')) {
				$product['subtract'] = trim($this->getCell($data, $i, $this->columns['subtract']['position'], 'true'));
			}
			if ($this->isEnabled('minimum')) {
				$product['minimum'] = trim($this->getCell($data, $i, $this->columns['minimum']['position'], '1'));
			}

			if ($this->isEnabled('date_added')) {
				$dateAdded = trim($this->getCell($data, $i, $this->columns['date_added']['position']));
				$product['date_added'] = ((is_string($dateAdded)) && (strlen($dateAdded) > 0)) ? $dateAdded : "NOW()";
			}
			if ($this->isEnabled('date_modified')) {
				$dateModified = trim($this->getCell($data, $i, $this->columns['date_modified']['position']));
				$product['date_modified'] = ((is_string($dateModified)) && (strlen($dateModified) > 0)) ? $dateModified : "NOW()";
			}
			if ($this->isEnabled('date_available')) {
				$dateAvailable = trim($this->getCell($data, $i, $this->columns['date_available']['position']));
				$product['date_available'] = ((is_string($dateAvailable)) && (strlen($dateAvailable) > 0)) ? $dateAvailable : "NOW()";
			}

			$products[$productId] = $product;
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

			if ($this->isEnabled('category_id') && $category['category_id']) {
				$categoryId = $category['category_id'];
			} else {
				continue;
			}

            $flag = false;
			$sql = "UPDATE `" . DB_PREFIX . "category` SET ";
			if ($this->isEnabled('parent_id') && isset($category['parent_id'])) {
				$sql .= "`parent_id`='" . (int) $category['parent_id'] . "',";
                $flag = true;
			}
			if ($this->isEnabled('image') && isset($category['image'])) {
				$sql .= "`image`='{$category['image']}', ";
                $flag = true;
			}
			if ($this->isEnabled('sort_order') && isset($category['sort_order'])) {
				$sql .= "`sort_order`='{$category['sort_order']}', ";
                $flag = true;
			}
			if ($this->isEnabled('date_added') && isset($category['date_added'])) {
				$category['date_added'] = ($category['date_added'] == 'NOW()') ? "{$category['date_added']}" : "'{$category['date_added']}'";
				$sql .= "`date_added`={$category['date_added']}, ";
                $flag = true;
			}
			if ($this->isEnabled('date_modified') && isset($category['date_modified'])) {
				$category['date_modified'] = ($category['date_modified'] == 'NOW()') ? "{$category['date_modified']}" : "'{$category['date_modified']}'";
				$sql .= "`date_modified`={$category['date_modified']}, ";
                $flag = true;
			}
			if ($this->isEnabled('status') && isset($category['status'])) {
				$status = ((strtolower($category['status']) == "true") || (strtoupper($category['status']) == "YES") || (strtoupper($category['status']) == "ENABLED")) ? 1 : 0;
				$sql .= "`status`='{$status}', ";
                $flag = true;
			}
			$sql = rtrim(trim($sql), ',') . " WHERE `category_id` = {$categoryId};";

            if ($flag) {
			    $this->db->query($sql);
            }

			if ($this->isEnabled('language_id') && $category['language_id']) {
				$languageId = $category['language_id'];
			} else {
                $languageId = $this->languageId;
			}


            if (isset($category['name']) || isset($category['description']) || isset($category['meta_description']) || isset($category['meta_keyword'])) {
                $sql = "UPDATE `" . DB_PREFIX . "category_description` SET ";
                if ($this->isEnabled('name') && isset($category['name'])) {
                    $sql .= "`name`='" . $this->db->escape($category['name']) . "',";
                }
                if ($this->isEnabled('description') && isset($category['description'])) {
                    $sql .= "`description`='" . $this->db->escape($category['description']) . "',";
                }
                if ($this->isEnabled('meta_description') && isset($category['meta_description'])) {
                    $sql .= "`meta_description`='" . $this->db->escape($category['meta_description']) . "',";
                }

                if ($this->isEnabled('meta_keyword') && isset($category['meta_keyword'])) {
                    $sql .= "`meta_keyword`='" . $this->db->escape($category['meta_keyword']) . "',";
                }

                $sql = rtrim(trim($sql), ',');
                $sql .= " WHERE `category_id` = {$categoryId}";
                $sql .= " AND `language_id` = {$languageId};";

                $this->db->query($sql);
            }


			if ($this->isEnabled('keyword') && isset($category['keyword'])) {
                $sql = "DELETE FROM " . DB_PREFIX . "url_alias  WHERE query = 'category_id={$categoryId}'";
                $this->db->query($sql);
                $sql = "INSERT INTO `" . DB_PREFIX . "url_alias` (`query`,`keyword`) ";
				$sql .= "VALUES ('category_id={$categoryId}','{$category['keyword']}') ";
				$sql .= "ON DUPLICATE KEY UPDATE `query`='category_id={$categoryId}',`keyword`='{$category['keyword']}';";
				$this->db->query($sql);
			}

			if ($this->isEnabled('store_ids') &&  isset($category['store_ids'])) {
				foreach ($category['store_ids'] as $storeId) {
					if ($storeId == '')
						continue;
                    $sql = "INSERT INTO `" . DB_PREFIX . "category_to_store` (`category_id`,`store_id`) ";
					$sql .= "VALUES ({$categoryId},{$storeId}) ";
					$sql .= "ON DUPLICATE KEY UPDATE `category_id`='{$categoryId}',`store_id`='{$storeId}';";
					$this->db->query($sql);
				}
			}
		}
		// final commit
		$this->db->query("COMMIT;");
		return true;
	}

    protected function buildFieldRelations(&$reader, $type)
    {
        $this->relation = array();
        $data = $reader->getSheet(0);
        $relation = array();
        switch($type) {
            case 'category':
                $config = $this->getCateforyFildsList();
                for ($i = 1; $i < count($config); $i++ ) {
                    $fieldName = $this->getCell($data, 0, $i);
                    if (isset($config[$fieldName]) ) {
                        $this->relation[$fieldName] = $config[$fieldName];
                        $this->relation[$fieldName]['position'] = $i;
                    }
                }
            break;
            case 'product':
                $config = $this->getProductFildsList();
                for ($i = 1; $i < count($config); $i++ ) {
                    $fieldName = $this->getCell($data, 0, $i);
                    if (isset($config[$fieldName]) ) {
                        $this->relation[$fieldName] = $config[$fieldName];
                        $this->relation[$fieldName]['position'] = $i;
                    }
                }
            break;
            default:
                return false;
            break;
        }
    }

	protected function uploadCategories(&$reader) {
        $this->columns = $this->getCateforyFildsList();
        $this->buildFieldRelations($reader, 'category');

		$data = $reader->getSheet(0);
		$categories = array();
		$isFirstRow = true;
		$i = 0;
		$j = 1;
		$k = $data->getHighestRow();

		for ($i = 0; $i < $k; $i+=1) {

			if ($isFirstRow) {
				$isFirstRow = false;
				continue;
			}
			$category = array();

            if (!isset($this->relation['category_id'])) {
                return;
            }

			$categoryId = trim($this->getCell($data, $i, $this->relation['category_id']['position']));
			if ($this->isEnabled('category_id') && $categoryId) {
				$category['category_id'] = $categoryId;
			} else {
				continue;
			}

			if ($this->isEnabled('parent_id') && isset($this->relation['parent_id'])) {
				$category['parent_id'] = trim($this->getCell($data, $i, $this->relation['parent_id']['position'], '0'));
			}

			if ($this->isEnabled('name') && isset($this->relation['name'])) {
				$name = trim($this->getCell($data, $i, $this->relation['name']['position']));
				$category['name'] = htmlentities($name, ENT_QUOTES, $this->detect_encoding($name));
			}

			if ($this->isEnabled('sort_order') && isset($this->relation['sort_order'])) {
				$category['sort_order'] = trim($this->getCell($data, $i, $this->relation['sort_order']['position']));
			}
			if ($this->isEnabled('image') && isset($this->relation['image'])) {
				$category['image'] = trim($this->getCell($data, $i, $this->relation['image']['position']));

			}
			if ($this->isEnabled('date_added') && isset($this->relation['date_added'])) {
				$dateAdded = trim($this->getCell($data, $i, $this->relation['date_added']['position']));
				$category['date_added'] = ((is_string($dateAdded)) && (strlen($dateAdded) > 0)) ? $dateAdded : "NOW()";
			}
			if ($this->isEnabled('date_modified') && isset($this->relation['date_modified'])) {
				$dateModified = trim($this->getCell($data, $i, $this->relation['date_modified']['position']));
				$category['date_modified'] = ((is_string($dateModified)) && (strlen($dateModified) > 0)) ? $dateModified : "NOW()";
			}
            /**
             * @todo Костыль детектед. Язык захардкожен, чтобы не выводить его в таблицу XLS
             */
			$langId = $this->languageId;//trim($this->getCell($data, $i, $j++, $this->languageId));
			if ($this->isEnabled('language_id') && $langId && ($langId == $this->languageId)) {
				$category['language_id'] = $langId;
			} else {
                $category['language_id'] = $this->languageId;
				//continue;
			}


			if ($this->isEnabled('keyword') && isset($this->relation['keyword'])) {
				$category['keyword'] = trim($this->getCell($data, $i, $this->relation['keyword']['position']));
			}

			if ($this->isEnabled('description') && isset($this->relation['description']) ) {
				$description = trim($this->getCell($data, $i, $this->relation['description']['position']));
				$category['description'] = htmlentities($description, ENT_QUOTES, $this->detect_encoding($description));
			}
			if ($this->isEnabled('meta_description') && isset($this->relation['meta_description'])) {
				$meta_description = trim($this->getCell($data, $i, $this->relation['meta_description']['position']));
				$category['meta_description'] = htmlentities($meta_description, ENT_QUOTES, $this->detect_encoding($meta_description));
			}
			if ($this->isEnabled('meta_keyword') && isset($this->relation['meta_keyword'])) {
				$meta_keyword = trim($this->getCell($data, $i, $this->relation['meta_keyword']['position']));
				$category['meta_keyword'] = htmlentities($meta_keyword, ENT_QUOTES, $this->detect_encoding($meta_keyword));
			}
			if ($this->isEnabled('store_ids') && isset($this->relation['store_ids'])) {
				$storeIds = $this->getCell($data, $i, $this->relation['store_ids']['position']);
				$storeIds = trim($this->clean($storeIds, false));
				$category['store_ids'] = ($storeIds == "") ? array() : explode(",", $storeIds);
				if ($category['store_ids'] === false) {
					$category['store_ids'] = array();
				}
			}
			if ($this->isEnabled('status') && isset($this->relation['status'])) {
				$category['status'] = trim($this->getCell($data, $i, $this->relation['status']['position'], 'true'));
			}

			$categories[$categoryId] = $category;
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

	function storeOptionDetailsIntoDatabase(&$options, &$optionIds) {



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


		foreach ($this->columns as $colId => $colData) {
			if (!$colData['enabled'])
				continue;
			$j++;
			$worksheet->setColumn($j, $j + 1, $colData['length'], $colData['format']);
			// The heading row
			$worksheet->writeString($i, $j - 1, ($colData['name'] ? $colData['name'] : $colId), $this->boxFormat);
		}

		$worksheet->setRow($i, 30, $this->boxFormat);


		$i++;
		$j = 0;
		$storeIds = $this->getStoreIdsForCategories();
		$query = "SELECT DISTINCT c.* , cd.*, ua.keyword FROM `" . DB_PREFIX . "category` c ";
		$query .= "INNER JOIN `" . DB_PREFIX . "category_description` cd ON cd.category_id = c.category_id ";
		$query .= " AND cd.language_id='{$this->languageId}' ";
		$query .= "LEFT JOIN `" . DB_PREFIX . "url_alias` ua ON ua.query=CONCAT('category_id=',c.category_id) ";
		$query .= "ORDER BY c.`parent_id`, `sort_order`, c.`category_id`;";

		$result = $this->db->query($query);

		foreach ($result->rows as $row) {
			if ($categoryId = (int) $row['category_id']) {
				$worksheet->write($i, $j++, $categoryId);
			} else {
				continue;
			}

			if ($this->isEnabled('parent_id'))
				$worksheet->write($i, $j++, $row['parent_id']);
			if ($this->isEnabled('name'))
				$worksheet->writeString($i, $j++, html_entity_decode($row['name'], ENT_QUOTES, 'UTF-8'));
			if ($this->isEnabled('sort_order'))
				$worksheet->write($i, $j++, $row['sort_order']);
			if ($this->isEnabled('image'))
				$worksheet->write($i, $j++, $row['image']);
			if ($this->isEnabled('date_added'))
				$worksheet->write($i, $j++, $row['date_added'], $this->textFormat);
			if ($this->isEnabled('date_modified'))
				$worksheet->write($i, $j++, $row['date_modified'], $this->textFormat);
			if ($this->isEnabled('language_id'))
				$worksheet->write($i, $j++, $row['language_id']);
			if ($this->isEnabled('keyword'))
				$worksheet->writeString($i, $j++, ($row['keyword']) ? $row['keyword'] : '' );
			if ($this->isEnabled('description'))
				$worksheet->writeString($i, $j++, html_entity_decode($row['description'], ENT_QUOTES, 'UTF-8'));
			if ($this->isEnabled('meta_description'))
				$worksheet->writeString($i, $j++, html_entity_decode($row['meta_description'], ENT_QUOTES, 'UTF-8'));
			if ($this->isEnabled('meta_keyword'))
				$worksheet->writeString($i, $j++, html_entity_decode($row['meta_keyword'], ENT_QUOTES, 'UTF-8'));
			if ($this->isEnabled('store_ids')) {
				$storeIdList = '';
				if (isset($storeIds[$categoryId])) {
					foreach ($storeIds[$categoryId] as $storeId) {
						$storeIdList .= ($storeIdList == '') ? $storeId : ',' . $storeId;
					}
				}
				$worksheet->write($i, $j++, $storeIdList, $this->textFormat);
			}
			if ($this->isEnabled('status'))
				$worksheet->write($i, $j++, ($row['status'] == 0) ? "false" : "true", $this->textFormat);

			$i++;
			$j = 0;
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
		// Set the column widths
        $config = $this->getConfig();
        $productFields = $config['product'];

        foreach ($this->columns as $key => $column ) {
            if ( !isset($productFields[$key]) ) {
                $this->columns[$key]['enabled'] = false;
            }
            else {
                $this->columns[$key]['enabled'] = true;
            }
        }

		$j = 0;
		$i = 0;
		foreach ($this->columns as $colId => $colData) {
			if ($colData['enabled'] && isset($productFields[$colId]) ) {
			    $j++;
			    $worksheet->setColumn($j, $j + 1, $colData['length'], $colData['format']);
			    // The heading row
			    $worksheet->writeString($i, $j - 1, ($colData['name'] ? $colData['name'] : $colId), $this->boxFormat);
            }
		}


		$worksheet->setRow($i, 30, $this->boxFormat);

		if ($this->isEnabled('images')) {
			// Get all additional product images
			$imageNames = array();
			$query = "SELECT DISTINCT ";
			$query .= "  p.product_id, ";
			$query .= "  pi.product_image_id AS image_id, ";
			$query .= "  pi.image AS filename ";
			$query .= "FROM `" . DB_PREFIX . "product` p ";
			$query .= "INNER JOIN `" . DB_PREFIX . "product_image` pi ON pi.product_id=p.product_id ";
			$query .= "ORDER BY product_id, image_id; ";
			$result = $this->db->query($query);
			foreach ($result->rows as $row) {
				$productId = $row['product_id'];
				$imageId = $row['image_id'];
				$imageName = $row['filename'];
				if (!isset($imageNames[$productId])) {
					$imageNames[$productId] = array();
					$imageNames[$productId][$imageId] = $imageName;
				} else {
					$imageNames[$productId][$imageId] = $imageName;
				}
			}
		}

		// The actual products data
		$i++;
		$j = 0;
		$storeIds = $this->getStoreIdsForProducts();
		$query = "SELECT ";
		$query .= "  p.product_id,";
		$query .= "  pd.name,";
		$query .= "  GROUP_CONCAT( DISTINCT CAST(pc.category_id AS CHAR(11)) SEPARATOR \",\" ) AS categories,";
		$query .= "  p.sku,";
		$query .= "  p.location,";
		$query .= "  p.quantity,";
		$query .= "  p.model,";
		$query .= "  m.name AS manufacturer,";
		$query .= "  p.image,";
		$query .= "  p.shipping,";
		$query .= "  p.price,";
		$query .= "  p.date_added,";
		$query .= "  p.date_modified,";
		$query .= "  p.date_available,";
		$query .= "  p.weight,";
		$query .= "  wc.unit,";
		$query .= "  p.length,";
		$query .= "  p.width,";
		$query .= "  p.height,";
		$query .= "  p.status,";
		$query .= "  p.tax_class_id,";
		$query .= "  p.viewed,";
		$query .= "  p.sort_order,";
		$query .= "  pd.language_id,";
		if ($this->isEnabled('keyword'))
			$query .= "  ua.keyword,";
		$query .= "  pd.description, ";
		$query .= "  pd.meta_description, ";
		$query .= "  pd.meta_keyword, ";
		$query .= "  p.stock_status_id, ";
		$query .= "  mc.unit AS length_unit, ";
		$query .= "  p.subtract, ";
		$query .= "  p.minimum, ";
		$query .= "  GROUP_CONCAT( DISTINCT CAST(pr.related_id AS CHAR(11)) SEPARATOR \",\" ) AS related, ";
		$query .= "  pd.tag AS tags ";
		$query .= "FROM `" . DB_PREFIX . "product` p ";
		$query .= "LEFT JOIN `" . DB_PREFIX . "product_description` pd ON p.product_id=pd.product_id ";
		$query .= "  AND pd.language_id='{$this->languageId}' ";
		$query .= "LEFT JOIN `" . DB_PREFIX . "product_to_category` pc ON p.product_id=pc.product_id ";
		if ($this->isEnabled('keyword'))
			$query .= "LEFT JOIN `" . DB_PREFIX . "url_alias` ua ON ua.query=CONCAT('product_id=',p.product_id) ";
		$query .= "LEFT JOIN `" . DB_PREFIX . "manufacturer` m ON m.manufacturer_id = p.manufacturer_id ";
		$query .= "LEFT JOIN `" . DB_PREFIX . "weight_class_description` wc ON wc.weight_class_id = p.weight_class_id ";
		$query .= "  AND wc.language_id='{$this->languageId}' ";
		$query .= "LEFT JOIN `" . DB_PREFIX . "length_class_description` mc ON mc.length_class_id=p.length_class_id ";
		$query .= "  AND mc.language_id='{$this->languageId}' ";
		$query .= "LEFT JOIN `" . DB_PREFIX . "product_related` pr ON pr.product_id=p.product_id ";
		$query .= "GROUP BY p.product_id ";
		$query .= "ORDER BY p.product_id, pc.category_id; ";

		$result = $this->db->query($query);
		foreach ($result->rows as $row) {
			$productId = $row['product_id'];
			if ($this->isEnabled('product_id')) {
				$worksheet->write($i, $j++, $productId);
			} else {
				continue;
			}
			if ($this->isEnabled('name'))
				$worksheet->writeString($i, $j++, html_entity_decode($row['name'], ENT_QUOTES, 'UTF-8'));
			if ($this->isEnabled('categories'))
				$worksheet->write($i, $j++, $row['categories'], $this->textFormat);
			if ($this->isEnabled('sku'))
				$worksheet->writeString($i, $j++, $row['sku']);
			if ($this->isEnabled('location'))
				$worksheet->writeString($i, $j++, $row['location']);
			if ($this->isEnabled('quantity'))
				$worksheet->write($i, $j++, $row['quantity']);
			if ($this->isEnabled('model'))
				$worksheet->writeString($i, $j++, $row['model']);
			if ($this->isEnabled('manufacturer'))
				$worksheet->writeString($i, $j++, $row['manufacturer']);
			if ($this->isEnabled('image'))
				$worksheet->writeString($i, $j++, $row['image']);
			if ($this->isEnabled('shipping'))
				$worksheet->write($i, $j++, ($row['shipping'] == 0) ? "no" : "yes", $this->textFormat);
			if ($this->isEnabled('price'))
				$worksheet->write($i, $j++, $row['price'], $this->priceFormat);
			if ($this->isEnabled('weight'))
				$worksheet->write($i, $j++, $row['weight'], $this->weightFormat);
			if ($this->isEnabled('unit'))
				$worksheet->writeString($i, $j++, $row['unit']);
			if ($this->isEnabled('length'))
				$worksheet->write($i, $j++, $row['length']);
			if ($this->isEnabled('width'))
				$worksheet->write($i, $j++, $row['width']);
			if ($this->isEnabled('height'))
				$worksheet->write($i, $j++, $row['height']);
			if ($this->isEnabled('length_unit'))
				$worksheet->writeString($i, $j++, $row['length_unit']);
			if ($this->isEnabled('status'))
				$worksheet->write($i, $j++, ($row['status'] == 0) ? "false" : "true", $this->textFormat);
			if ($this->isEnabled('tax_class_id'))
				$worksheet->write($i, $j++, $row['tax_class_id']);
			if ($this->isEnabled('viewed'))
				$worksheet->write($i, $j++, $row['viewed']);
			if ($this->isEnabled('language_id'))
				$worksheet->write($i, $j++, $row['language_id']);
			if ($this->isEnabled('keyword'))
				$worksheet->writeString($i, $j++, ($row['keyword']) ? $row['keyword'] : '' );
			if ($this->isEnabled('description'))
				$worksheet->writeString($i, $j++, html_entity_decode($row['description'], ENT_QUOTES, 'UTF-8'), $this->textFormat, true);
			if ($this->isEnabled('meta_description'))
				$worksheet->write($i, $j++, html_entity_decode($row['meta_description'], ENT_QUOTES, 'UTF-8'), $this->textFormat);
			if ($this->isEnabled('meta_keyword'))
				$worksheet->write($i, $j++, html_entity_decode($row['meta_keyword'], ENT_QUOTES, 'UTF-8'), $this->textFormat);
			if ($this->isEnabled('images')) {
				$names = "";
				if (isset($imageNames[$productId])) {
					$first = true;
					foreach ($imageNames[$productId] AS $name) {
						if (!$first) {
							$names .= ",\n";
						}
						$first = false;
						$names .= $name;
					}
				}
				$worksheet->write($i, $j++, $names, $this->textFormat);
			}
			if ($this->isEnabled('stock_status_id'))
				$worksheet->write($i, $j++, $row['stock_status_id']);
			if ($this->isEnabled('store_ids')) {
				$storeIdList = '';
				if (isset($storeIds[$productId])) {
					foreach ($storeIds[$productId] as $storeId) {
						$storeIdList .= ($storeIdList == '') ? $storeId : ',' . $storeId;
					}
				}
				$worksheet->write($i, $j++, $storeIdList, $this->textFormat);
			}
			if ($this->isEnabled('related'))
				$worksheet->write($i, $j++, $row['related'], $this->textFormat);
			if ($this->isEnabled('tags'))
				$worksheet->write($i, $j++, $row['tags'], $this->textFormat);
			if ($this->isEnabled('sort_order'))
				$worksheet->write($i, $j++, $row['sort_order']);
			if ($this->isEnabled('subtract'))
				$worksheet->write($i, $j++, ($row['subtract'] == 0) ? "false" : "true", $this->textFormat);
			if ($this->isEnabled('minimum'))
				$worksheet->write($i, $j++, $row['minimum']);
			if ($this->isEnabled('date_added'))
				$worksheet->write($i, $j++, $row['date_added'], $this->textFormat);
			if ($this->isEnabled('date_modified'))
				$worksheet->write($i, $j++, $row['date_modified'], $this->textFormat);
			if ($this->isEnabled('date_available'))
				$worksheet->write($i, $j++, $row['date_available'], $this->textFormat);
			$i++;
			$j = 0;
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
