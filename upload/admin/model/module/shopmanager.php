<?php

ini_set('include_path', ini_get('include_path') . ':' . DIR_SYSTEM . 'library/Shopmanager/');
ini_set("max_execution_time", 360);

static $config = NULL;
static $log = NULL;

// Error Handlers
function shopmanager_error_handler($errno, $errstr, $errfile, $errline)
{
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

/**
 * create error handler
 */
function shopmanager_error_shutdown_handler()
{
    $last_error = error_get_last();
    if ($last_error['type'] === E_ERROR) {
        shopmanager_error_handler(E_ERROR, $last_error['message'], $last_error['file'], $last_error['line']);
    }
}

/**
 * ModelModuleShopmanager
 *
 * Processed all interactions of controller  with database and business logic
 *
 * @copyright  2012 Stfalcon (http://stfalcon.com/)
 */
class ModelModuleShopmanager extends Model
{

    /**
     * @var string : database config table (without prefix)
     */
    public $configTable = 'shopmanager_config';

    /**
     * @var array | export/import settings
     */
    public $settings = array();

    /**
     * @var array | link between reading XLS column and  field in $column
     */
    public $relation = array();

    /**
     * Current language
     * @var int current language id
     */
    protected $languageId = 1;

    /**
     * Cell's format
     *
     * @var Spreadsheet_Excel_Writer format
     */
    protected $priceFormat, $boxFormat, $weightFormat, $textFormat;

    /**
     * @var array | columns which used on import export
     */
    protected $columns = array();

    /**
     * Get config data from base
     *
     * @return array config data
     */
    public function getConfig()
    {
        $sql = "SELECT * FROM " . DB_PREFIX . $this->configTable;

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

    public function install(){
         $sql = "CREATE TABLE  IF NOT EXISTS `". DB_PREFIX ."shopmanager_config` (
            `config_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
            `group` varchar(60) NOT NULL,
            `field` varchar(255) NOT NULL,
            PRIMARY KEY (`config_id`),
            KEY `group` (`group`(10))
        ) ENGINE=MyISAM DEFAULT CHARSET=utf8;";
         $this->db->query($sql);
    }

    /**
     * Write config data to base
     *
     * @param string $group group of config data ('category', 'product')
     * @param array $data parameters for save (List of checked fields)
     */
    public function setConfig($group, $data)
    {
        $fieldList = array();
        //get possible data from the reference array
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
                    $sql = substr($sql, 0, -1);
                    $this->db->query($sql .= ';');
                }
                break;
            default:
                break;
        }
    }


    /**
     *  Return all fields which used in category export import
     *
     * format:
     * array_key - unique field name
     * name - string | user friendly name
     * length - int | length of field on Excel file
     * format - list | data format from module var ($priceFormat, $boxFormat, $weightFormat, $textFormat)
     * enabled - boolean | availability of current field (true by default)
     * multirow - boolean | true if field has multilanguage data,
     *
     * @return array : list of category used fields
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
     *  Return all fields which used in product export import
     *
     * format:
     * array_key - unique field name
     * name - string | user friendly name
     * length - int | length of field on Excel file
     * format - list | data format from module var ($priceFormat, $boxFormat, $weightFormat, $textFormat)
     * enabled - boolean | availability of current field (true by default)
     * multirow - boolean | true if field has multilanguage data,
     *
     * @return array : list of category used fields
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
    }

    /**
     * @var array table structure from database (for value by default)
     */
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
            'product_id' => '0',
            'language_id' => '0',
            'name' => '',
            'description' => '',
            'meta_description' => '',
            'meta_keyword' => '',
            'tag' => '',
        ),
    );

    /**
     * Loading XLS data to MySQL
     *
     * @param string $filename
     * @param string $group ('category', 'product')
     *
     * @return boolean is successful finished
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

        switch (trim($group)) {
            case 'category':
                //clearing the cache
                $this->cache->delete('category');
                $this->cache->delete('category_description');
                $this->cache->delete('url_alias');
                //processed data
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
     * Load group config data
     *
     * @param $group | string processed group type
     * @return bool processing success
     */
    public function loadColumns($group)
    {
        switch ($group) {
            case 'category':
                $this->columns = $this->getCateforyFildsList();
                break;
            case 'product':
                $this->columns = $this->getProductFildsList();
                break;
            default:
                return false;
        }
        return true;
    }

    /**
     * Export group to Excel
     *
     * @param string $group
     * @return boolean processing success
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

        $this->priceFormat = & $workbook->addFormat(
            array('Size' => 10,
                'Align' => 'right',
                'NumFormat' => '######0.00')
        );
        $this->boxFormat = & $workbook->addFormat(
            array('Size' => 10,
                'vAlign' =>
                'vequal_space')
        );
        $this->weightFormat = & $workbook->addFormat(
            array('Size' => 10,
                'Align' => 'right',
                'NumFormat' => '##0.00')
        );
        $this->textFormat = & $workbook->addFormat(
            array('Size' => 10,
                'NumFormat' => "@")
        );

        $group = strtolower(trim($group));

        // sending HTTP headers
        $workbook->send('backup_' . $group . '_' . date('Ymd') . '.xls');

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
            default:
                return false;
        }

        $worksheet->freezePanes(array(1, 1, 1, 1));

        $workbook->close();

        exit;
    }

    /**
     * Initialisation of module
     */
    protected function init()
    {
        global $config;
        global $log;
        $config = $this->config;
        $log = $this->log;
        set_error_handler('shopmanager_error_handler', E_ALL);
        register_shutdown_function('shopmanager_error_shutdown_handler');
        $this->loadDefaultLanguage();
    }

    /**
     * Clearing import data
     *
     * @param string $str
     * @param bool $allowBlanks
     * @return string
     */
    protected function clean($str, $allowBlanks = false)
    {
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


    /**
     * execute query from array ( paste together rows )
     *
     * @param array $sql
     *
     */
    protected function import($sql)
    {
        foreach (explode(";\n", $sql) as $sql) {
            $sql = trim($sql);
            if ($sql) {
                $this->db->query($sql);
            }
        }
    }


    /**
     * get system default weight element
     *
     * @return mixed
     */
    protected function getDefaultWeightUnit()
    {
        return $this->config->get('config_weight_class');
    }

    /**
     * get system default dimension element
     *
     * @return mixed
     */
    protected function getDefaultMeasurementUnit()
    {
        return $this->config->get('config_length_class');
    }

    /**
     * create new manufacturer from imported list if notexists
     *
     * @param array $products products data structure
     * @return bool operation success
     */
    protected function updateManufacturersInDB(&$products)
    {
        // find all manufacturers already stored in the database
        $keyArray = array_keys($products);
        $manufacturerName = array();
        if (isset($products[$keyArray[0]]['manufacturer'])) {
            foreach ($products as $product) {
                if (trim($product['manufacturer']) != '') {
                    $manufacturerName[$product['manufacturer']] = $product['manufacturer'];
                }
            }
            $sql = "SELECT * FROM `" . DB_PREFIX . "manufacturer`;";
            $result = $this->db->query($sql);

            $query = '';
            foreach ($manufacturerName as $mName) {
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

                    $query .= "('" . implode($manufacturerSimple, "','") . "')";
                    $query = str_replace("'NULL'", 'NULL', $query);

                    $sql = "INSERT INTO `" . DB_PREFIX . "manufacturer` VALUES {$query}";
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
                        $query .= "('" . implode(array($lastId, $store_id), "','") . "'),";
                    }

                    $query = rtrim(trim($query), ',');
                    $sql = "INSERT INTO `" . DB_PREFIX . "manufacturer_to_store` (`manufacturer_id`,`store_id`) VALUES {$query} ;";
                    $this->db->query($sql);
                }
            }
        } else {
            return true;
        }
        return true;
    }

    /**
     * Get weight class data
     *
     * @return array all weight classes
     */
    protected function getWeightClassIds()
    {
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

    /**
     * Get length class data
     *
     * @return array all length classes
     */
    protected function getLengthClassIds()
    {
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

    /**
     * @param string $name manufacturer name
     * @return int manufacturer id
     */
    protected function getManufacturerByName($name)
    {
        $defaultManufacturer = 0;

        $result = $this->db->query("SELECT `manufacturer_id` FROM `" . DB_PREFIX . "manufacturer` WHERE name = '{$name}'; ");
        $data = $result->row;

        if (isset($data['manufacturer_id'])) {
            return $data['manufacturer_id'];
        } else {
            return $defaultManufacturer;
        }
    }

    /**
     * insert product data to database
     *
     * @param array $products products list
     * @return bool operation success
     */
    protected function updateProductsInDB($products)
    {
        $this->import("START TRANSACTION;\n");


        // store or update manufacturers
        if (!$this->updateManufacturersInDB($products)) {
            $this->db->query('ROLLBACK;');
            return false;
        }

        // get weight classes
        $weightClassIds = $this->getWeightClassIds();

        // get length classes
        $lengthClassIds = $this->getLengthClassIds();

        // generate and execute SQL for storing the products
        foreach ($products as $key => $product) {
            //get old product data
            $productData = array();
            if (trim($product['product_id']) != '') {
                $sql = "SELECT * FROM `" . DB_PREFIX . "product` WHERE product_id = '{$product['product_id']}'; ";
                $result = $this->db->query($sql);
                $productData = $result->row;
            }

            //merge old data with simple data
            $productData = array_merge($this->tableSimple['product'], $productData);
            $insert = '';

            // fill newly reading data to product data array
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
                if (isset($weightClassIds[$product['weight']])) {
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
                if (isset($weightClassIds[$product['length_unit']])) {
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

            //formed sql query from data
            $query = "('" . implode($productData, "','") . "')";
            $query = str_replace("'NULL'", 'NULL', $query);

            $insert = rtrim(trim($insert), ',');

            if ($insert != '') {
                $sql = "INSERT INTO `" . DB_PREFIX . "product` VALUES {$query} ON DUPLICATE KEY UPDATE {$insert} ; ";
                $this->db->query($sql);
            }

            $lastId = trim($product['product_id']);
            if ($lastId == '') {
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

            foreach ($languageList as $code => $language) {
                $sql = "SELECT * FROM `" . DB_PREFIX . "product_description` WHERE product_id = '{$lastId}' AND language_id = '{$language['language_id']}' ;";
                $result = $this->db->query($sql);

                $productDescription = $this->tableSimple['product_description'];
                $productDescription = array_merge($productDescription, $result->row);

                $flag = false;
                $productDescription['product_id'] = $lastId;
                $productDescription['language_id'] = $language['language_id'];

                if (isset($product['name']) && isset($product['name'][$language['language_id']])) {
                    $productDescription['name'] = $product['name'][$language['language_id']];
                    $flag = true;
                }
                if (isset($product['description']) && isset($product['description'][$language['language_id']])) {
                    $productDescription['description'] = $product['description'][$language['language_id']];
                    $flag = true;
                }
                if (isset($product['meta_description']) && isset($product['meta_description'][$language['language_id']])) {
                    $productDescription['meta_description'] = $product['meta_description'][$language['language_id']];
                    $flag = true;
                }
                if (isset($product['meta_keyword']) && isset($product['meta_keyword'][$language['language_id']])) {
                    $productDescription['meta_keyword'] = $product['meta_keyword'][$language['language_id']];
                    $flag = true;
                }

                if ($flag) {
                    $sql = "DELETE FROM `" . DB_PREFIX . "product_description` WHERE product_id = '{$lastId}' AND language_id = '{$language['language_id']}' ;";
                    $this->db->query($sql);

                    foreach ($productDescription as $fieldName => $productData) {
                        $productDescription[$fieldName] = str_replace('\'', '&#39;', $productData);
                    }

                    $query = "('" . implode($productDescription, "','") . "')";
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
     * Detecting message ebncoding
     *
     * @param string $str
     * @return string
     */
    protected function detect_encoding($str)
    {
        return mb_detect_encoding($str, 'UTF-8,ISO-8859-15,ISO-8859-1,cp1251,KOI8-R');
    }

    /**
     * read XlS file, processed data and transfer to DB insert
     *
     * @param Spreadsheet_Excel_Reader $reader
     * @return bool success
     */
    function uploadProducts(&$reader)
    {

        //read first Sheet
        $data = $reader->getSheet(0);
        //check if correct file
        if ($data->getTitle() != 'Product') {
            $this->log('Wrong file type');
            return false;
        }

        //get columns list
        $this->columns = $this->getProductFildsList();
        //weed empty columns
        $this->buildFieldRelations($reader, 'product');
        $this->columns = $this->relation;

        //get default units
        $defaultWeightUnit = $this->getDefaultWeightUnit();
        $defaultMeasurementUnit = $this->getDefaultMeasurementUnit();
        $defaultStockStatusId = $this->config->get('config_stock_status_id');

        //get range of Excel product list
        $products = array();
        $isFirstRow = true;
        $i = 0;
        $j = 1;
        $k = $data->getHighestRow();


        //reading data, form data array
        for ($i = 2; $i < $k; $i += 1) {
            $product = array();
            $productId = trim($this->getCell($data, $i, $j++));

            foreach ($this->relation as $relKey => $relation) {
                if (isset($relation['position'])) {
                    $tmp = trim($this->getCell($data, $i, $relation['position']));
                    //$product[$relKey] = htmlentities($tmp, ENT_QUOTES, $this->detect_encoding($tmp));
                    $product[$relKey] = $this->encodeCharacters($tmp);
                } else {
                    foreach ($relation as $langCode => $sublanguage) {
                        $tmp = trim($this->getCell($data, $i, $relation[$langCode]['position']));
                        //$product[$relKey][$langCode] = htmlentities($tmp, ENT_QUOTES, $this->detect_encoding($tmp));
                        $product[$relKey][$langCode] = $this->encodeCharacters($tmp);
                    }
                }
            }

            $products[] = $product;
            $j = 1;
        }

        // transfer to database insert
        return $this->updateProductsInDB($products);
    }

    protected function encodeCharacters ($message) {
        return str_replace('\'','&#39',$message);
    }

    /**
     * update categories in database
     *
     * @param array $categories categories list
     * @return bool operation success
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
            if (isset($category['category_id']) && trim($category['category_id']) != '') {
                $sql = "SELECT * FROM " . DB_PREFIX . "category WHERE category_id = {$category['category_id']}";
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

            $query = "INSERT INTO `" . DB_PREFIX . "category` VALUES {$sql} ON DUPLICATE KEY UPDATE {$insert} ";
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

                $sql = "SELECT * FROM `" . DB_PREFIX . "category_description` WHERE category_id = {$lastId} and language_id = {$language['language_id']} LIMIT 1; ";
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

                $query = "DELETE FROM `" . DB_PREFIX . "category_description` WHERE category_id = '{$catDescSimple['category_id']}' AND language_id = '{$catDescSimple['language_id']}';";
                $this->db->query($query);

                $query = "INSERT INTO `" . DB_PREFIX . "category_description` VALUES {$sql} ";
                $this->db->query($query);
            }

            //adding keywords to  url_alias;
            if (isset($category['keyword']) && trim($category['keyword']) != '') {
                $sql = "DELETE FROM " . DB_PREFIX . "url_alias  WHERE query = 'category_id={$lastId}'";
                $this->db->query($sql);

                $sql = "INSERT INTO `" . DB_PREFIX . "url_alias` (`query`,`keyword`) ";
                $sql .= "VALUES ('category_id={$lastId}','{$category['keyword']}') ";
                $this->db->query($sql);
            }
        }
        // final commit
        $this->db->query("COMMIT;");
        return true;
    }

    /**
     * create link between category settings array and columns in XLS file
     *
     * @param Spreadsheet_Excel_reader $reader
     * @param string $type data type
     * @return bool operation success
     */
    protected function buildFieldRelations(&$reader, $type)
    {
        $this->load->model('localisation/language');
        $languageList = $this->model_localisation_language->getLanguages();

        $this->relation = array();
        $data = $reader->getSheet(0);
        $relation = array();
        switch ($type) {
            case 'category':
                $config = $this->getCateforyFildsList();
                $colCount = PHPExcel_Cell::columnIndexFromString($data->getHighestColumn());

                for ($i = 1; $i < $colCount + 1; $i++) {
                    $fieldName = $this->getCell($data, 0, $i);
                    foreach ($config as $key => $configField) {
                        if ($fieldName == $key) {
                            $this->relation[$fieldName] = $config[$fieldName];
                            $this->relation[$fieldName]['position'] = $i;
                            break;
                        } else if (preg_match("/^{$key}[_a-z0-9]{0,3}$/", $fieldName)) {
                            $langArray = preg_split("/^{$key}[_]{1}/", $fieldName);

                            foreach ($languageList as $language) {
                                if ($language['code'] == $langArray[1] ){
                                    $languageCode = $language['language_id'];
                                    break;
                                }
                            }

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

                for ($i = 1; $i < $colCount + 1; $i++) {
                    $fieldName = $this->getCell($data, 0, $i);
                    foreach ($config as $key => $configField) {
                        if ($fieldName == $key) {
                            $this->relation[$fieldName] = $config[$fieldName];
                            $this->relation[$fieldName]['position'] = $i;
                            break;
                        } else if (preg_match("/^{$key}[_a-z0-9]{0,3}$/", $fieldName)) {
                            $langArray = preg_split("/^{$key}[_]{1}/", $fieldName);

                            foreach ($languageList as $language) {
                                if ($language['code'] == $langArray[1] ){
                                    $languageCode = $language['language_id'];
                                    break;
                                }
                            }

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

    /**
     * processed category data from XLS file and sending to DB insert
     *
     * @param Spreadsheet_Excel_reader $reader
     * @return bool operation success
     */
    protected function uploadCategories(&$reader)
    {
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

        for ($i = 1; $i < $k; $i += 1) {
            if ($isFirstRow) {
                $isFirstRow = false;
                continue;
            }
            $category = array();
            if (!isset($this->relation['category_id'])) {
                return;
            }

            foreach ($this->relation as $relKey => $relation) {
                if (isset($relation['position'])) {
                    $tmp = trim($this->getCell($data, $i, $relation['position']));
                    //$category[$relKey] = htmlentities($tmp, ENT_QUOTES, $this->detect_encoding($tmp));
                    $category[$relKey] = $tmp;
                } else {
                    foreach ($relation as $langCode => $sublanguage) {
                        $tmp = trim($this->getCell($data, $i, $relation[$langCode]['position']));
                        $category[$relKey][$langCode] = htmlentities($tmp, ENT_QUOTES, $this->detect_encoding($tmp));
                        $category[$relKey][$langCode] = $tmp;
                    }
                }
            }

            $categories[] = $category;
            $j = 1;
        }

        return $this->updateCategoriesInDB($categories);
    }

    function storeAdditionalImagesIntoDatabase(&$reader)
    {
// start transaction
        $sql = "START TRANSACTION;\n";

// delete old additional product images from database
        $sql = "DELETE FROM `" . DB_PREFIX . "product_image`";
        $this->db->query($sql);

// insert new additional product images into database
        $data = & $reader->getSheet(1); // Products worksheet
        $maxImageId = 0;

        $k = $data->getHighestRow();
        for ($i = 1; $i < $k; $i += 1) {
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

    function uploadImages(&$reader)
    {
        $ok = $this->storeAdditionalImagesIntoDatabase($reader);
        return $ok;
    }

    /**
     * Get table cell data from reader
     *
     * @param Spreadsheet_excel_reader_sheet $worksheet
     * @param int $row
     * @param int $col
     * @param string $default_val
     * @return string Cell data
     */
    protected function getCell(&$worksheet, $row, $col, $default_val = '')
    {
        $col -= 1; // we use 1-based, PHPExcel uses 0-based column index
        $row += 1; // we use 0-based, PHPExcel used 1-based row index
        return ($worksheet->cellExistsByColumnAndRow($col, $row)) ? $worksheet->getCellByColumnAndRow($col, $row)->getValue() : $default_val;
    }

    /**
     * Get all stores linked with categories
     *
     * @return array category store IDs
     */
    protected function getStoreIdsForCategories()
    {
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

    /**
     * read category data from database and push to XLS
     *
     * @param Spreadsheet_excel_writer $worksheet
     * @return bool sucess
     */
    function populateCategoriesWorksheet(&$worksheet)
    {
        // Set the column widths
        $j = 0;
        $i = 0;

        $config = $this->getConfig();
        $categoryFields = isset($config['category']) ? $config['category'] : array();

        //switch off unselected categories
        foreach ($this->columns as $key => $column) {
            if (!isset($categoryFields[$key])) {
                $this->columns[$key]['enabled'] = false;
            }
        }

        $this->load->model('localisation/language');
        $languageList = $this->model_localisation_language->getLanguages();

        //formed SQL query
        $query = "SELECT DISTINCT c.* , ua.keyword";
        foreach ($languageList as $code => $language) {
            $query .= ", description_{$language['code']}.name AS name_{$language['code']}";
            $query .= ", description_{$language['code']}.description AS description_{$language['code']}";
            $query .= ", description_{$language['code']}.meta_description AS meta_description_{$language['code']}";
            $query .= ", description_{$language['code']}.meta_keyword AS meta_keyword_{$language['code']} ";
        }
        $query .= "FROM `" . DB_PREFIX . "category` AS c ";

        foreach ($languageList as $code => $language) {
            $query .= "INNER JOIN `" . DB_PREFIX . "category_description` AS description_{$language['code']} ON description_{$language['code']}.category_id = c.category_id  AND description_{$language['code']}.language_id='{$language['language_id']}' ";
        }

        $query .= "LEFT JOIN `" . DB_PREFIX . "url_alias` ua ON ua.query=CONCAT('category_id=',c.category_id) ";
        $query .= "ORDER BY c.`parent_id`, `sort_order`, c.`category_id`;";

        $result = $this->db->query($query);

        if (sizeof($result->row) == 0) {
            return false;
        }

        //create data array with enclosure for multilanguage columns
        $writeColumns = array();
        foreach ($result->row as $key => $row) {
            foreach ($this->columns as $cname => $column) {
                if (preg_match("/^{$cname}[_a-z0-9]{0,3}$/", $key) && $column['enabled']) {
                    if (isset($column['multirow']) && $column['multirow']) {
                        foreach ($languageList as $code => $language) {
                            $writeColumns[$cname . "_{$language['code']}"] = array(
                                'name' => "{$column['name']}({$language['code']})",
                                'length' => $column['length'],
                                'format' => $column['format'],
                            );
                        }
                    } else {
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
            $worksheet->writeString($i + 1, $j - 1, ($colData['name'] ? $colData['name'] : $colId), $this->boxFormat);
        }
        $i++;
        $worksheet->setRow(0, 1, $this->boxFormat); // set first row height 1 px (ID fields)
        $worksheet->setRow($i, 30, $this->boxFormat);
        $i++;
        $j = 0;
        $storeIds = $this->getStoreIdsForCategories();

        // write data
        foreach ($result->rows as $row) {
            foreach ($writeColumns as $colName => $col) {
                $worksheet->write($i, $col['position'], $row[$colName]);
            }
            $i++;
        }
    }

    /**
     * get linked store data for products
     *
     * @return array
     */
    function getStoreIdsForProducts()
    {
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


    /**
     * read product data from database and push to XLS
     *
     * @param Spreadsheet_excel_writer $worksheet
     * @return bool sucess
     */
    function populateProductsWorksheet(&$worksheet)
    {
        $this->load->model('localisation/language');
        $languageList = $this->model_localisation_language->getLanguages();

        // Set the column widths
        $config = $this->getConfig();
        $productFields = $config['product'];

        $i = 1;
        foreach ($this->columns as $key => $column) {
            if (!isset($productFields[$key])) {
                $this->columns[$key]['enabled'] = false;
            } else {
                $this->columns[$key]['enabled'] = true;
                $i++;
            }
        }

        $j = 1;
        $i = 0;
        $colCount = 0;
        $writeColumns = array();
        foreach ($this->columns as $colId => $colData) {
            if ($colData['enabled']) {
                if (!$colData['multirow']) {
                    $worksheet->setColumn($j, $j + 1, $colData['length'], $colData['format']);
                    $worksheet->writeString($i, $j - 1, $colId, $this->boxFormat);
                    $worksheet->writeString($i + 1, $j - 1, ($colData['name'] ? $colData['name'] : $colId), $this->boxFormat);
                    $writeColumns[$colId] = $colData;
                    $writeColumns[$colId]['position'] = $j - 1;
                    $colCount++;
                    $j++;
                } else {
                    foreach ($languageList as $code => $language) {
                        $worksheet->setColumn($j, $j + 1, $colData['length'], $colData['format']);
                        $worksheet->writeString($i, $j - 1, "{$colId}_{$language['code']}", $this->boxFormat);
                        $worksheet->writeString($i + 1, $j - 1, "{$colData['name']}({$language['code']})", $this->boxFormat);
                        $writeColumns["{$colId}_{$language['code']}"] = $colData;
                        $writeColumns["{$colId}_{$language['code']}"]['position'] = $j - 1;
                        $colCount++;
                        $j++;
                    }
                }
            }
        }

        $worksheet->setRow($i, 1, $this->boxFormat);
        $worksheet->setRow($i + 1, 30, $this->boxFormat);

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
            $query .= "_pd_{$language['code']}.description AS description_{$language['code']}, ";
            $query .= "_pd_{$language['code']}.name AS name_{$language['code']}, ";
            $query .= "_pd_{$language['code']}.meta_description AS meta_description_{$language['code']}, ";
            $query .= "_pd_{$language['code']}.meta_keyword AS meta_keyword_{$language['code']}, ";
            $query .= "_pd_{$language['code']}.tag AS tags_{$language['code']}, ";
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
            $query .= "LEFT JOIN `" . DB_PREFIX . "product_description` _pd_{$language['code']} ON _pd_{$language['code']}.product_id=_p.product_id AND _pd_{$language['code']}.language_id = '{$language['language_id']}' ";
        }
        ;
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
                if (isset($writeColumns[$cellName])) {
                    $worksheet->write($i, $writeColumns[$cellName]['position'], $cell, $this->textFormat);
                }
            }
        }
    }

    /**
     * Load default language
     *
     * @return type
     */
    protected function loadDefaultLanguage()
    {
        $code = $this->config->get('config_language');
        $sql = "SELECT language_id FROM `" . DB_PREFIX . "language` WHERE code = '$code'";
        $result = $this->db->query($sql);
        $this->languageId = ($result->row && $result->row['language_id']) ? $result->row['language_id'] : 1;
    }

    /**
     * Log writer
     * @param type $text
     */
    protected function log($text)
    {
        if ($text = trim($text)) {
            error_log(date('Y-m-d H:i:s - ', time()) . "Export/Import: {$text}\n", 3, DIR_LOGS . "error.txt");
        }
    }
}