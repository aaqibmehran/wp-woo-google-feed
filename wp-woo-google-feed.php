<?php

/**
 * Plugin Name: WooCommerce Product Excel Importer
 * Plugin URI: https://github.com/aaqibmehran/wp-woo-google-feed
 * Description: Fetches WooCommerce products and imports them with metadata and creates an Excel file in the project base URL every 5 hours.
 * Version: 1.0.0
 * Author: Aaqib Mehran
 * Author URI: https://github.com/aaqibmehran
 * Text Domain: woocommerce-product-excel-importer
 * License: MIT
 */

// Define the plugin directory.
define('WCPE_PLUGIN_DIR', plugin_dir_path(__FILE__));

require_once plugin_dir_path(__FILE__) . 'vendor/autoload.php';

// Use the `PhpOffice\PhpSpreadsheet\Spreadsheet` class.
use PhpOffice\PhpSpreadsheet\Spreadsheet;

// Register a cron job to fetch and import products every 5 hours.
add_action('wp_loaded', function () {
    if (!wp_next_scheduled('wcpe_fetch_and_import_products')) {
        wp_schedule_event(time() + 5 * 18000, 'every_5_hours', 'wcpe_fetch_and_import_products');
    }
});

// Fetch and import products.
function wcpe_fetch_and_import_products()
{
    // Get all WooCommerce products.
    $products = wc_get_products();

    // Create a new `Spreadsheet` object.
    $excel = new Spreadsheet();

    // Set the worksheet title.
    $excel->getActiveSheet()->setTitle('WooCommerce Products');

    // Write the header row.
    $excel->getActiveSheet()->setCellValueByColumnAndRow(0, 1, 'Product ID');
    $excel->getActiveSheet()->setCellValueByColumnAndRow(1, 1, 'Product Title');
    $excel->getActiveSheet()->setCellValueByColumnAndRow(2, 1, 'Product Description');
    $excel->getActiveSheet()->setCellValueByColumnAndRow(3, 1, 'Product Price');
    $excel->getActiveSheet()->setCellValueByColumnAndRow(4, 1, 'Product Metadata');

    // Write the product data to the Excel file.
    foreach ($products as $product) {
        $excel->getActiveSheet()->setCellValueByColumnAndRow(0, $product->get_id() + 1, $product->get_id());
        $excel->getActiveSheet()->setCellValueByColumnAndRow(1, $product->get_id() + 1, $product->get_title());
        $excel->getActiveSheet()->setCellValueByColumnAndRow(2, $product->get_id() + 1, $product->get_description());
        $excel->getActiveSheet()->setCellValueByColumnAndRow(3, $product->get_id() + 1, $product->get_price());
        $excel->getActiveSheet()->setCellValueByColumnAndRow(4, $product->get_id() + 1, serialize($product->get_meta_data()));
    }

    // Save the Excel file to the project base URL.
    $writer = new PhpOffice\PhpSpreadsheet\Writer\Xlsx($excel);
    $writer->save(get_site_url() . '/woocommerce-products.xlsx');
}

// Activate the plugin.
register_activation_hook(__FILE__, 'wcpe_activate');
function wcpe_activate()
{
    // Schedule the cron job to fetch and import products.
    wp_schedule_event(time() + 5 * 18000, 'every_5_hours', 'wcpe_fetch_and_import_products');
}

// Deactivate the plugin.
register_deactivation_hook(__FILE__, 'wcpe_deactivate');
function wcpe_deactivate()
{
    // Clear the cron job.
    wp_clear_scheduled_hook('wcpe_fetch_and_import_products');
}
