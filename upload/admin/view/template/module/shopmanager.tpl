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
        <div id="tabs" class="htabs">
            <a href="#tab-category" style="display:block;">Category</a>
            <a href="#tab-product" style="display:block;">Product</a>
        </div>
        <!-- Category task -->
        <div id="tab-category">
            <table class="form">
                <tr>
                    <td>Импорт:</td>
                    <td>Экспорт:</td>
                </tr>
                <tr>
                    <td style="width: 50%;vertical-align: top;">
                        <form action="<?php echo $action; ?>" method="post" enctype="multipart/form-data" id="form_category">
                            <input type="file" name="upload" />
                            <input type="hidden" name="group" value="category"/>
                            <a onclick="$('#form_category').submit();" class="button"><span><?php echo $button_import; ?></span></a>
                        </form>
                    </td>
                    <td>
                        <form action="<?php echo $export; ?>" method="post" id="form_category_export">
                            <input type="hidden" name="group" value="category"/>
                            <input type="hidden" name="test" value="1"/>
                            <span class="roll-link">Additional settings</span>
                            <div style="display: none;" class="scrollbox-container">
                                <div class="scrollbox" style="margin-bottom: 5px; disply:none;">
                                    <?php $class = 'odd'; ?>
                                    <?php foreach ($category_fields as $key => $data) : ?>
                                    <?php $class = ($class == 'even' ? 'odd' : 'even'); ?>
                                    <div class="<?php echo $class; ?>">
                                        <input type="checkbox" name="category[]" value="<?php echo $key; ?>" <?php echo isset($selected_categories[$key]) ? 'checked="checked"' : '';?> />
                                        <?php echo $key; ?></div>
                                    <?php endforeach; ?>
                                </div>
                                <a onclick="$(this).parent().find(':checkbox').attr('checked', true);"><?php echo $text_select_all; ?></a> / <a onclick="$(this).parent().find(':checkbox').attr('checked', false);"><?php echo $text_unselect_all; ?></a>
                            </div>
                            <br><br>
                            <a onclick="$('#form_category_export').submit();" class="button"><span><?php echo $button_export; ?></span></a>
                            <!--a onclick="location='<?php echo $export; ?>&group=category'" class="button"><span><?php echo $button_export; ?></span></a-->
                        </form>
                    </td>
                </tr>
            </table>
        </div>
        <!-- Products task -->
        <div id="tab-product">
            <table class="form">
                <tr>
                    <td>Импорт:</td>
                    <td>Экспорт:</td>
                </tr>
                <tr>
                    <td style="width: 50%;vertical-align: top;">
                        <form action="<?php echo $action; ?>" method="post" enctype="multipart/form-data" id="form_product">
                            <input type="file" name="upload" />
                            <input type="hidden" name="group" value="product"/>
                            <a onclick="$('#form_product').submit();" class="button"><span><?php echo $button_import; ?></span></a>
                        </form>
                    </td>
                    <td>
                        <form action="<?php echo $export; ?>" method="post" enctype="multipart/form-data" id="form_product_export">
                            <input type="hidden" name="group" value="product"/>
                            <input type="hidden" name="test" value="1"/>
                            <span class="roll-link">Additional settings</span>
                            <div style="display: none;" onclick="turnObject(this);false;" class="scrollbox-container">
                                <div class="scrollbox" style="margin-bottom: 5px;">
                                    <?php $class = 'odd'; ?>
                                    <?php foreach ($product_fields as $key => $field) : ?>
                                    <?php $class = ($class == 'even' ? 'odd' : 'even'); ?>
                                    <div class="<?php echo $class; ?>">
                                        <input type="checkbox" name="product[]" value="<?php echo $key; ?>" <?php echo isset($selected_products[$key]) ? 'checked="checked"' : '';?> />
                                        <?php echo $key; ?></div>
                                    <?php endforeach; ?>
                                </div>
                                <a onclick="$(this).parent().find(':checkbox').attr('checked', true);"><?php echo $text_select_all; ?></a> / <a onclick="$(this).parent().find(':checkbox').attr('checked', false);"><?php echo $text_unselect_all; ?></a>
                            </div>
                            <br><br>
                            <!--a onclick="location='<?php echo $export; ?>&group=product'" class="button"><span><?php echo $button_export; ?></span></a-->
                            <a onclick="$('#form_product_export').submit();" class="button"><span><?php echo $button_export; ?></span></a>
                        </form>
                    </td>
                </tr>
            </table>
        </div>
    </div>
</div>
<script type="text/javascript"><!--
$('#tabs a').tabs();
$('#languages a').tabs();
$('#vtab-option a').tabs();

$(document).ready(function(){
    $('div.scrollbox-container').each(function(index, element){
        elem = $(element);
        if ($('input[type="checkbox"]:checked', elem)[0] == undefined ) {
            elem.css('display', 'block');
        }
    });

    $('span.roll-link').click(function( event ){
        element = $(event.target);
        $('div.scrollbox-container', element.parent()).toggle();
    });
});

//--></script>
<?php echo $footer; ?>