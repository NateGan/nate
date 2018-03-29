<?php
include_once 'admin-order-data.php';
$list_table = new Dobot_Order_List();
$list_table->prepare_items();
$allOrderCount = $list_table::record_count(true);
$statusArr = wc_get_order_statuses();
unset($statusArr['wc-failed']);
?>
    <div class="wrap">
        <h1><?php echo esc_html(get_admin_page_title()); ?></h1>
        <div class="notice notice-warning">
            <p>导出须知：</p>
            <p>1、先勾选要订单，然后在导出</p>
            <p>2、提供了2种表格数据格式（csv格式、excel格式）</p>
            <p>3、只能导出当前页面的订单；可以根据列表中筛选，导出指定的订单</p>
            <p>4、导出数据中，每个订单的各个产品信息已合并成一列,格式：<br/></p>
            <p>&nbsp;&nbsp;&nbsp;&nbsp;(产品1) * (数量1) * (价格1);&nbsp;&nbsp;(产品2) * (数量2) * (价格2);<br/></p>
        </div>
        <!-- Forms are NOT created automatically, so you need to wrap the table in one to use features like bulk actions -->
		<?php $current = !isset($_GET['post_status']) ? 'class="current"' : ''; ?>
        <ul class="subsubsub">
            <li class="all">
                <a href="<?php echo admin_url('admin.php?page=wc-export-order') ?>"
                    <?php echo $current ?>>
                    全部<span class="count"><?php echo $allOrderCount?></span>
                </a> |
            </li>
			<?php $i = 1;foreach ($statusArr as $value => $status): ?>
				<?php
                if($value != 'wc-failed')
				$current = '';
				if (isset($_GET['post_status'])): ?>
					<?php
					if ($_REQUEST['post_status'] == $value) {
						$current = 'class="current" ';
					} else {
						$current = '';
					}
					?>
				<?php endif; ?>
                <li class="<?php echo $value ?>">
                    <a href="<?php echo admin_url('admin.php?page=wc-export-order') ?>&post_status=<?php echo $value ?>" <?php echo $current ?>><?php echo $status ?>
                        <span class="count">(<?php echo wc_orders_count(str_replace('wc-', '', $value)) ?>)</span></a>
					<?php
					if ($i < count($statusArr)): ?> | <?php endif;?>
                </li>
				<?php $i++;endforeach; ?>
        </ul>
        <form id="contest-collect-filter" method="get">
            <input type="hidden" name="post_status" value="<?php echo isset($_REQUEST['post_status']) ? $_REQUEST['post_status'] : ''?>">
            <!-- For plugins, we also need to ensure that the form posts back to our current page -->
            <input type="hidden" name="page" value="<?php echo $_REQUEST['page'] ?>"/>
            <!-- Now we can render the completed list table -->
			<?php $list_table->display() ?>
        </form>
    </div>
    <script type="text/javascript">
       if( jQuery('div.notice-warning').length > 1){
           jQuery('div.notice-warning').eq(0).remove();
       }
    </script>
<?php