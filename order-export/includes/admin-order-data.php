<?php
/**
 * Created by PhpStorm.
 * User: silk-nate
 * Date: 2018/3/23
 * Time: 11:47
 */
if (!class_exists('WP_List_Table')) {
	require_once(ABSPATH . 'wp-admin/includes/class-wp-list-table.php');
}

try {
	require_once(dirname(__DIR__) . '/PHPExcel/Classes/PHPExcel.php');
	require_once(dirname(__DIR__) . '/PHPExcel/Classes/PHPExcel/Writer/Excel2007.php');
} catch (Exception $e) {
	die($e->getMessage());
}

class Dobot_Order_List extends WP_List_Table
{
	
	/**
	 * current post type
	 *
	 * @var string
	 */
	protected $post_type = 'shop_order';
	
	/**
	 * 导出表格头
	 */
	const export_headerLine = array(
		'order_title' => '订单号',
		'transaction_id' => '支付ID',
		'order_date' => '订单时间',
		'order_status' => '订单状态',
		'shipping_name' => '姓名',
		'shipping_address' => '地址',
		'shipping_phone' => '电话',
		'shipping_email' => '邮箱',
		'buy_product_info' => '购买产品',
		'order_total' => '订单总额',
		'discount' => '优惠金额',
		'shipping_total' => '运费',
		'note' => '备注信息'
	);
	
	public function __construct($args = array())
	{
		parent::__construct([
			'singular' => __('dobot_export_order', 'dobot'), //singular name of the listed records
			'plural' => __('dobot_export_orders', 'dobot'), //plural name of the listed records
			'ajax' => false //should this table support ajax?
		]);
	}
	
	/**
	 * sql 条件组建
	 *
	 * @param $sql
	 * @return string
	 */
	public static function buildWhere($sql)
	{
		if (isset($_REQUEST['m']) && $_REQUEST['m']) {
			$date = $_REQUEST['m'];
			$year = substr($date, 0, 4);
			$moth = substr($date, 4, 2);
			$time = mktime(20, 20, 20, 2, 1, $year);//取得一个日期的 Unix 时间戳;
			if (date("t", $time) == 29) { //格式化时间，并且判断2月是否是29天；
				$run = true;
			} else {
				$run = false;
			}
			if ((int)$moth == 2 && (!$run)) {
				$mothDays = 28;
			} else if ((int)$moth == 2 && ($run)) {
				$mothDays = 29;
			} else if (in_array((int)$moth, array(1, 3, 5, 7, 8, 10, 12))) {
				$mothDays = 31;
			} else {
				$mothDays = 30;
			}
			$mothFirst = date('Y-m-d H:i:s', strtotime($year . $moth . '01'));
			
			$mothLast = date('Y-m-d H:i:s', strtotime($year . $moth . $mothDays));
			
			$sql .= ' and post_date >= "' . $mothFirst . '" AND post_date <= "' . $mothLast . '" ';
		}
		if (isset($_REQUEST['post_status'])) {
			if ($_REQUEST['post_status']) {
				$status = $_REQUEST['post_status'];
				$sql .= ' and post_status = "' . $status . '" ';
			}
		}
		
		return $sql;
	}
	
	/**
	 * 获取订单记录
	 *
	 * @param int $per_page
	 * @param int $page_number
	 * @return array|null|object
	 */
	public static function get_orders($per_page = 15, $page_number = 1)
	{
		global $wpdb;
		
		$sql = "select ID,post_date FROM {$wpdb->prefix}posts WHERE post_type = 'shop_order'";
		
		$sql = self::buildWhere($sql);
		
		if (!empty($_REQUEST['orderby'])) {
			$sql .= ' ORDER BY ' . esc_sql($_REQUEST['orderby']);
			
			$sql .= !empty($_REQUEST['order']) ? ' ' . esc_sql($_REQUEST['order']) : ' ASC';
		}
		$sql .= " LIMIT $per_page";
		
		$sql .= ' OFFSET ' . ($page_number - 1) * $per_page;
		
		$result = $wpdb->get_results($sql, 'ARRAY_A');
		
		return $result;
	}
	
	/**
	 * Returns the count of records in the database.
	 *
	 * @param  $flag bool
	 * @return null|string
	 */
	public static function record_count($flag = false)
	{
		global $wpdb;
		
		$sql = "SELECT COUNT(*) FROM {$wpdb->prefix}posts WHERE post_type = 'shop_order'";
		
		if ($flag === false) {
			$sql = self::buildWhere($sql);
		}
		
		return $wpdb->get_var($sql);
	}
	
	public function no_items()
	{
		_e('没有找到任何记录。', 'dobot');
	}
	
	/**
	 * @param $orderId
	 * @return WC_Order
	 */
	public static function get_order($orderId)
	{
		return wc_get_order($orderId);
	}
	
	/**
	 * @param $orderId
	 * @return string
	 */
	public static function getOrderIncrementId($orderId)
	{
		$order = self::get_order($orderId);
		$link = admin_url('post.php') . "?post=" . $orderId . "&action=edit";
		return "<a href='" . $link . "' target='blank' title='" . $order->get_order_number() . "'>#" . $order->get_order_number() . "</a>";
	}
	
	/***
	 * @param object $item
	 * @param string $column_name
	 */
	public function column_default($item, $column_name)
	{
		$post = get_post($item['ID']);
		$the_order = self::get_order($item['ID']);
		switch ($column_name) {
			case 'order_title':
				echo self::getOrderIncrementId($item['ID']);
				break;
			case 'order_items' :
				
				echo '<a href="#" class="show_order_items">' . apply_filters('woocommerce_admin_order_item_count', sprintf(_n('%d item', '%d items', $the_order->get_item_count(), 'woocommerce'), $the_order->get_item_count()), $the_order) . '</a>';
				
				if (sizeof($the_order->get_items()) > 0) {
					
					echo '<table class="order_items" cellspacing="0">';
					
					foreach ($the_order->get_items() as $item) {
						$product = apply_filters('woocommerce_order_item_product', $the_order->get_product_from_item($item), $item);
						$item_meta = new WC_Order_Item_Meta($item, $product);
						$item_meta_html = $item_meta->display(true, true);
						?>
                        <tr class="<?php echo apply_filters('woocommerce_admin_order_item_class', '', $item, $the_order); ?>">
                            <td class="qty"><?php echo absint($item['qty']); ?></td>
                            <td class="name">
								<?php if ($product) : ?>
									<?php echo (wc_product_sku_enabled() && $product->get_sku()) ? $product->get_sku() . ' - ' : ''; ?>
                                    <a href="<?php echo get_edit_post_link($product->id); ?>"
                                       title="<?php echo apply_filters('woocommerce_order_item_name', $item['name'], $item, false); ?>"><?php echo apply_filters('woocommerce_order_item_name', $item['name'], $item, false); ?></a>
								<?php else : ?>
									<?php echo apply_filters('woocommerce_order_item_name', $item['name'], $item, false); ?>
								<?php endif; ?>
								<?php if (!empty($item_meta_html)) : ?>
									<?php echo wc_help_tip($item_meta_html); ?>
								<?php endif; ?>
                            </td>
                        </tr>
						<?php
					}
					
					echo '</table>';
					
				} else echo '&ndash;';
				break;
			case 'order_date':
				if ('0000-00-00 00:00:00' == $post->post_date) {
					$t_time = $h_time = __('Unpublished', 'woocommerce');
				} else {
					$t_time = get_the_time(__('Y/m/d g:i:s A', 'woocommerce'), $post);
					$h_time = get_the_time(__('Y/m/d', 'woocommerce'), $post);
				}
				
				echo '<abbr title="' . esc_attr($t_time) . '">' . esc_html(apply_filters('post_date_column_time', $h_time, $post)) . '</abbr>';
				
				break;
			case 'order_status':
				printf('<mark class="%s tips" data-tip="%s">%s</mark>', sanitize_title($the_order->get_status()), wc_get_order_status_name($the_order->get_status()), wc_get_order_status_name($the_order->get_status()));
				break;
			case 'shipping_name':
				echo $the_order->get_formatted_shipping_full_name();
				break;
			case 'shipping_address':
				if ($address = $the_order->get_formatted_shipping_address()) {
					echo '<a target="_blank" href="' . esc_url($the_order->get_shipping_address_map_url()) . '">' . esc_html(preg_replace('#<br\s*/?>#i', ', ', $address)) . '</a>';
				} else {
					echo '&ndash;';
				}
				break;
			case 'shipping_phone':
				echo esc_html($the_order->billing_phone);
				break;
			case 'shipping_email':
				echo '<a href="' . esc_url('mailto:' . $the_order->billing_email) . '">' . esc_html($the_order->billing_email) . '</a>';
				break;
			case 'order_total':
				echo $the_order->get_formatted_order_total();
				if ($the_order->payment_method_title) {
					echo '<small class="meta">' . __('Via', 'woocommerce') . ' ' . esc_html($the_order->payment_method_title) . '</small>';
				}
				break;
			default:
				echo '未知'; //Show the whole array for troubleshooting purposes
				break;
		}
	}
	
	/**
	 * Render the bulk edit checkbox
	 *
	 * @param array $item
	 *
	 * @return string
	 */
	public function column_cb($item)
	{
		return sprintf(
			'<input type="checkbox" name="order_id[]" value="%s" />', $item['ID']
		);
	}
	
	/**
	 *  Associative array of columns
	 *
	 * @return array
	 */
	public function get_columns()
	{
		$columns = [
			'cb' => '<input type="checkbox" />',
			'order_status' => '<span class="status_head tips" data-tip="' . esc_attr__('Status', 'woocommerce') . '">' . esc_attr__('Status', 'woocommerce') . '</span>',
			'order_title' => __('#订单', 'dobot'),
			'order_items' => __('物品', 'dobot'),
			'shipping_name' => __('姓名', 'dobot'),
			'shipping_address' => __('地址', 'dobot'),
			//'shipping_phone'    => __('电话', 'dobot'),
			'shipping_email' => __('邮箱', 'dobot'),
			'order_date' => __('日期', 'dobot'),
			'order_total' => __('总额', 'dobot'),
		];
		
		return $columns;
	}
	
	
	public function get_bulk_actions()
	{
		$actions = array(
			'export_csv' => '导出Csv',
			'export_excel' => '导出Excel',
		);
		return $actions;
	}
	
	/**
	 * 每页显示订单的个数工具栏
	 */
	public function pager_show_list()
	{
		$count = self::record_count(true);//获取订单总数
		$per_gage = array(
			$count => '所有订单',
			'10' => '10个',
			'20' => '20个',
			'40' => '40个',
			'60' => '60个',
			'80' => '80个',
			'100' => '100个',
		);
		?>
        <label for="filter-by-date" class="screen-reader-text"><?php echo '每页显示订单数量'; ?></label>
        <select name="limit_show" id="per_page_show" title="每页显示订单数量">
			<?php foreach ($per_gage as $page => $label): ?>
				<?php
				if (isset($_GET['limit_show']) && $_GET['limit_show']) {
					$limit_Show = $_GET['limit_show'];
				} else {
					$limit_Show = 10;
				}
				$selected = $page == $limit_Show ? 'selected = "selected"' : ''
				?>
                <option value="<?php echo $page ?>" <?php echo $selected ?>><?php echo $label ?></option>
			<?php endforeach; ?>
        </select>
		<?php
	}
	
	/**
	 * 表格左边的顶部筛选
	 *
	 * @param string $which
	 */
	protected function extra_tablenav($which)
	{
		?>
        <div class="alignleft actions">
			<?php
			if ('top' === $which && !is_singular()) {
				ob_start();
				$this->months_dropdown($this->post_type);
				$this->pager_show_list();
				$output = ob_get_clean();
				
				if (!empty($output)) {
					echo $output;
					submit_button(__('Filter'), '', 'filter_action', false, array('id' => 'post-query-submit'));
				}
			}
			?>
        </div>
		<?php
	}
	
	/**
	 * 批量操作数据
	 */
	public function process_bulk_action()
	{
		$dataFileObj = null;
		if (in_array($this->current_action(), array('export_csv', 'export_excel'))) {
			$ids = isset($_REQUEST['order_id']) ? $_REQUEST['order_id'] : array();
			if (!$ids) {
				die("<script>alert('请先选择要导出的订单！');window.history.back(-1);</script> ;");
			}
			if ($this->current_action() === 'export_excel') {
				self::exportExcel($ids);
			} else if ($this->current_action() === 'export_csv') {
				self::exportCsv($ids);
			}
		}
	}
	
	/**
	 *
	 * 导出excel格式的数据
	 *
	 * @param $ids
	 */
	public static function exportExcel($ids)
	{
		$result[0] = self::export_headerLine;
		foreach ($ids as $key => $id) {//导出程序
			$result[($key + 1)] = self::get_order_export_info($id);
		}
		$objPHPExcel = new PHPExcel();
		$objPHPExcel->getProperties()->setTitle("Dobot Order Export");
		$objPHPExcel->getProperties()->setSubject("Dobot Order Export");
		$obj_Sheet = $objPHPExcel->getActiveSheet();
		//$obj_Sheet->getDefaultRowDimension()->setRowHeight(15);
		$obj_Sheet->setTitle('s1');//设置当前活动Sheet名称
		$obj_Sheet->getDefaultStyle()->getAlignment()//设置居中显示
		->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER)
			->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
		//$obj_Sheet->getStyle("I2:I".count($result))->getAlignment()->setWrapText(TRUE);//自动换行
		$obj_Sheet->fromArray($result);//填充数组
		$obj_Writer = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');//创建工厂对象
		//操作1 保存文件
		$filePath = dirname(__DIR__) . '/order-export/order-export.xlsx';
		$obj_Writer->save($filePath);
		$fileZipPath = self::zipFile($filePath);
		self::downloadFile($fileZipPath);
		unlink(dirname(__DIR__) . '/order-export/order-export.xlsx');
	}
	
	/**
	 * 导出csv格式的数据
	 * @param $ids
	 */
	public static function exportCsv($ids)
	{
		$filePath = dirname(__DIR__) . '/order-export/orders_export.csv';
		$handle = fopen($filePath, 'w+');
		//文件头部
		fputcsv($handle, self::export_headerLine);
		foreach ($ids as $id) {//导出程序
			$data = self::get_order_export_info($id);
			fputcsv($handle, array_values($data));
		}
		fclose($handle);
		$fileZipPath = self::zipFile($filePath);
		self::downloadFile($fileZipPath);
		unlink(dirname(__DIR__) . '/order-export/orders_export.csv');
	}
	
	
	public static function zipFile($filePath)
	{
		//先压缩
		$date = date('Ymd');
		$zip = new ZipArchive();
		$zip->open(dirname(__DIR__) . '/order-export/order-export-' . $date . '.zip', ZipArchive::CREATE);   //打开压缩包
		$zip->addFile($filePath, basename($filePath));   //向压缩包中添加文件
		$zip->close();
		$zipPath = dirname(__DIR__) . '/order-export/order-export-' . $date . '.zip';
		return $zipPath;
	}
	
	/**
	 * 文件下载
	 *
	 * @param $filePath
	 */
	public static function downloadFile($filePath)
	{
		ob_start();
		if (file_exists($filePath)) {
			$fileinfo = pathinfo($filePath);
			$filename = $fileinfo['basename']; //文件名
			header("Content-type:application/x-" . $fileinfo['extension']);
			header("Content-Disposition:attachment;filename = " . $filename);
			header("Accept-ranges:bytes");
			header("Accept-length:" . filesize($fileinfo));
			readfile($filePath);
			//下载成功删除文件
			unlink($filePath);
		} else {
			echo "<script>alert('下载文件不存在或生成文件时失败。')</script>";
		}
		return true;
	}
	
	/**
	 * 导出订单信息
	 *
	 * @param $order_id
	 * @return array
	 */
	public static function get_order_export_info($order_id)
	{
		$order = self::get_order($order_id);
		$incrementNumber = $order->get_order_number();
		$order_status = wc_get_order_status_name($order->get_status());
		$shippingName = $order->get_formatted_shipping_full_name();
		$phone = $order->billing_phone;
		$email = $order->billing_email;
		$address = preg_replace('#<br\s*/?>#i', ', ', $order->get_formatted_shipping_address());
		$buy_product_info = '';
		foreach ($order->get_items() as $item_id => $item) {
			$price = esc_attr(wc_format_localized_price($item['line_subtotal']));
			$buy_product_info .= "(" . $item['name'] . ')*(' . $item['qty'] . ')*(' . $price . ");\r\n";
		}
		$buy_product_info = trim($buy_product_info, ";\r\n");
		return [
			'order_title' => $incrementNumber,
			'transaction_id' => $order->get_transaction_id(),
			'order_date' => get_the_time(__('Y/m/d', 'woocommerce'), $order_id),
			'order_status' => $order_status,
			'shipping_name' => $shippingName,
			'shipping_address' => $address,
			'shipping_phone' => $phone,
			'shipping_email' => $email,
			'buy_product_info' => $buy_product_info,
			'order_total' => esc_attr(wc_format_localized_price($order->get_total())),
			'discount' => esc_attr(wc_format_localized_price($order->get_total_discount())),
			'shipping_total' => $order->get_total_shipping() ? esc_attr(wc_format_localized_price($order->get_total_shipping())) : 0,
			'note' => join('<br/>', $order->get_customer_order_notes()),
		];
	}
	
	
	/**
	 * 表格排序字段
	 *
	 * @return array
	 */
	public function get_sortable_columns()
	{
		$sortable_columns = array(
			'order_title' => array('ID', true),
			'order_date' => array('post_date', false),
		);
		return $sortable_columns;
	}
	
	
	/**
	 * 表格数据
	 *
	 * Handles data query and filter, sorting, and pagination.
	 */
	public function prepare_items()
	{
		$current_page = $_GET['paged'] ? $_GET['paged'] : 1;
		
		$per_page = 10;
		if (isset($_REQUEST['limit_show'])) {
			if ($_REQUEST['limit_show']) {
				$per_page = $_REQUEST['limit_show'];
			}
		}
		
		$columns = $this->get_columns();
		
		$hidden = array();
		
		$sortable = $this->get_sortable_columns();
		
		$this->process_bulk_action();
		
		$this->_column_headers = array($columns, $hidden, $sortable);
		
		$total_items = self::record_count();
		
		$this->items = self::get_orders($per_page, $current_page);
		
		$this->set_pagination_args([
			'total_items' => $total_items, //WE have to calculate the total number of items
			'per_page' => $per_page,//WE have to determine how many items to show on a page
			'total_pages' => ceil($total_items / $per_page)
		]);
	}
}