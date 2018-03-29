<?php
/**
 * Created by PhpStorm.
 * User: silk-nate
 * Date: 2018/3/29
 * Time: 13:26
 */


/*
Plugin Name: 后台订单导出
Plugin URI:
Description: 提供后台订单导出功能
Version: 1.0.1
Author: Nate
Author URI:
License:
*/


if ( ! defined( 'ABSPATH' ) ) {
	exit; // Exit if accessed directly
}
class Dobot_Admin_OrderExport
{
	static public function output()
	{
		include_once ('includes/admin-order-data.php');
		include_once( 'includes/html-admin-page-order.php' );
	}
}

class Order_Export
{
	public function __construct()
	{
		add_action( 'admin_menu', array(&$this, 'order_export_menu' ));
		add_action( 'admin_enqueue_scripts', array( $this, 'admin_styles_dobot' ) );
		add_action( 'admin_enqueue_scripts', array( $this, 'admin_scripts_dobot' ) );
	}
	
	public function admin_styles_dobot()
	{
		wp_enqueue_style( 'woocommerce_admin_styles', WC()->plugin_url() . '/assets/css/admin.css', array(), WC_VERSION );
	}
	
	public function admin_scripts_dobot()
	{
		$suffix  = defined( 'SCRIPT_DEBUG' ) && SCRIPT_DEBUG ? '' : '.min';
		wp_register_script( 'woocommerce_admin', WC()->plugin_url() . '/assets/js/admin/woocommerce_admin' . $suffix . '.js', array( 'jquery', 'jquery-blockui', 'jquery-ui-sortable', 'jquery-ui-widget', 'jquery-ui-core', 'jquery-tiptip' ), '2.70');
		wp_register_script( 'jquery-blockui', WC()->plugin_url() . '/assets/js/jquery-blockui/jquery.blockUI' . $suffix . '.js', array( 'jquery' ), '2.70', true );
		wp_register_script( 'jquery-tiptip', WC()->plugin_url() . '/assets/js/jquery-tiptip/jquery.tipTip' . $suffix . '.js', array( 'jquery' ), '2.70', true );
		wp_enqueue_script('woocommerce_admin');
		wp_enqueue_script('jquery-blockui');
		wp_enqueue_script('jquery-tiptip');
	}
	
	public function order_export_menu()
	{
		add_submenu_page( 'woocommerce', __( '导出订单', 'woocommerce' ),  __( '导出订单', 'woocommerce' ) , 'view_woocommerce_reports', 'wc-export-order', array( $this, 'export_order_page' ) );
	}
	
	public function export_order_page()
	{
		Dobot_Admin_OrderExport::output();
	}
	
}
return new Order_Export();

