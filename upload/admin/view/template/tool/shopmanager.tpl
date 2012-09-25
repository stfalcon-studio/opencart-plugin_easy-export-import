<?php echo $header; ?>
<?php if ($error_warning) { ?>
<div class="warning"><?php echo $error_warning; ?></div>
<?php } ?>
<?php if ($success) { ?>
<div class="success"><?php echo $success; ?></div>
<?php } ?>
<div class="box">
	<div class="left"></div>
	<div class="right"></div>
	<div class="heading">
		<h1>
            <img src="view/image/backup.png">
            <?php echo $heading_title; ?>
        </h1>
	</div>
	<div class="content">
		<table class="form">
			<!-- Category task -->
			<tr>
				<td><?php echo $entry_category_task; ?>:</td>
				<td>
					<form action="<?php echo $action; ?>" method="post" enctype="multipart/form-data" id="form_category">
						<input type="file" name="upload" />
						<input type="hidden" name="group" value="category"/>
					</form>
				</td>
				<td>
					<a onclick="$('#form_category').submit();" class="button"><span><?php echo $button_import; ?></span></a>
					<a onclick="location='<?php echo $export; ?>&group=category'" class="button"><span><?php echo $button_export; ?></span></a>
				</td>
			</tr>
			<!-- Products task -->
			<tr>
				<td><?php echo $entry_product_task; ?>:</td>
				<td>
					<form action="<?php echo $action; ?>" method="post" enctype="multipart/form-data" id="form_product">
						<input type="file" name="upload" />
						<input type="hidden" name="group" value="product"/>
					</form>
				</td>
				<td>
					<a onclick="$('#form_product').submit();" class="button"><span><?php echo $button_import; ?></span></a>
					<a onclick="location='<?php echo $export; ?>&group=product'" class="button"><span><?php echo $button_export; ?></span></a>
				</td>
			</tr>
		</table>

	</div>
</div>
<?php echo $footer; ?>