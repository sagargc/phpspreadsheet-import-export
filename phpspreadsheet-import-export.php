<?php
/**
 * The plugin import and export data into database
 * 
 * This file is read by WordPress to generate the plugin information in the plugin 
 * admin area. This file also includes all of the dependencies used by the plugin, 
 * registers the activation and deactivation functions, and defines a function 
 * that starts the plugin.
 * 
 * @link			https://github.com/sagargc/phpspreadsheet-import-export
 * @since			1.0.0
 * @package			PhpSpreadsheet_Import_Export
 * 
 * Plugin Name: PhpSpreadsheet Import and Export
 * Plugin URI: https://github.com/sagargc/phpspreadsheet-import-export
 * Description: PhpSpreadsheet Import and Export Plugin for WordPress
 * Version: 1.0.0
 * Author: Sagar G C
 * Author URI: http://sagargc.com.np 
 * License: GPL-2.0+
 * License URI: http://www.gnu.org/licenses/gpl-2.0.html  
 * Text Domain: phpspreadsheet-import-export
 * Domain Path: /languages 
 */

define("PHP_SPREADSHEET_PLUGIN_URL", WP_PLUGIN_URL.'/'.basename(dirname(__FILE__)));
define("PHP_SPREADSHEET_PLUGIN_DIR", WP_PLUGIN_DIR.'/'.basename(dirname(__FILE__)));
require_once (get_template_directory() .'/vendor/autoload.php'); 
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Csv;
use PhpOffice\PhpSpreadsheet\Reader\Xls;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Calculation\TextData\Replace;

if ( ! defined( 'ABSPATH' ) ) {
	exit; // Exit if accessed directly.
}
// Check if PHP version is below 8.0.
if(phpversion() < 8.0){
    add_action( 'admin_notices', 'php_spreadsheet_admin_notice' );
    function php_spreadsheet_admin_notice() {
        ?>
        <div class="notice notice-error is-dismissible">
            <p><?php _e( 'PHP Spreadsheet Plugin requires PHP 8.0 or higher.', 'php_spreadsheet' ); ?></p>
        </div>
        <?php
    }
    return;
}
// Add menus.
add_action('admin_menu', 'my_alumnis_add_import_menu');
function my_alumnis_add_import_menu(){
	add_submenu_page( 'edit.php?post_type=alumni', 'Add alumni', 'Add alumni','manage_options', 'admin.php?page=alumni_add','alumni_add');
    add_submenu_page( 'edit.php?post_type=alumni',  'Alumni Import', 'Alumni Import', 'manage_options', 'admin.php?page=alumnis_import_upload', 'alumnis_import_upload');
    add_menu_page ( 'Export Users', 'Export User Type', 'manage_options', 'phpspreadsheet-dashboard', 'phpspreadsheet_dashboard', 'dashicons-media-spreadsheet
	', 10 );	
}

// Add alumni.
function alumni_add() { 
    ?>
    <div class="wrap">
        <h1><?php _e( 'Add new alumni', 'textdomain' ); ?></h1>
       <!--  <p><?php _e( 'Upload xls-file to import alumni', 'textdomain' ); ?></p>
        <div class="import">Import alumnis: <input type="file" name="import" /> <button type="submit">Import</button></div>
        <hr> -->
        <p><?php _e( 'All alumni information is required', 'textdomain' ); ?></p>
        <form action="" id="myform">
        	<div class="error"></div>
        	<label>First Name *</label><br>
        	<input type="text" name="alumni_first_name" id="alumni_first_name" /><br>
        	<label>Last Name *</label><br>
        	<input type="text" name="alumni_last_name" /><br>
        	<label>Email *</label><br>
        	<input type="text" name="alumni_email" /><br><br>
        	<input type="button" id='save_alumni' value="Save Alumni" name="Save">  
        </form>
        <!-- <hr>
        <h3>Alumni List</h3> -->
        <style>
        	.error { color:#f00; }
        </style>
        <script src="https://cdn.jsdelivr.net/npm/jquery-validation@1.17.0/dist/jquery.validate.min.js"></script>
        <script>
        	jQuery(document).ready(function(){
        		jQuery('#save_alumni').click(function(){
        			jQuery.validator.setDefaults({
  						debug: true,
  						success: "valid"
						});
        		     jQuery("#myform").validate({
        		     	rules: {
						    alumni_first_name: {
						      required: true
						    },
						    alumni_last_name: {
						      required: true
						    },
						    alumni_email: {
						      required: true,
						      email: true
						    }
						  }

        		     });
        		    var formStatus=jQuery('#myform').validate().form();
 					if(true == formStatus){
        		     
        			var formData=new FormData(document.getElementById('myform')); 
        			formData.append("action", "add_admin_alumni"); 
        			 
        			jQuery.ajax({
        				type:"post",
        				url:ajaxurl,
        				data: formData,
        				cache: false,
            			processData: false, 
            			contentType: false,  
        				success: function(response) {
        					if(jQuery.trim(response) == 'Success')
        					{
                                alert( jQuery.trim( response ) );
        						document.location.href = '<?php echo get_site_url().'/wp-admin/edit.php?post_type=alumni'; ?>';
        					}
        					else
        					{
        						alert(response);
        					}        					
        				}
        			});

        		}
        		});
        	});
        </script>
<?php 
           /* $args = array( 'post_type' => 'alumni' );

			$postlist = new WP_Query( $args );

			if( $postlist->have_posts() ) {
				?>
				<table>
					<thead>
						<tr><th>Alumni name</th><th>Email</th>
						</tr>
					</thead>
					<tbody>
				<?php
				while( $postlist->have_posts() ) {
						$postlist->the_post();
						$post_ID=get_the_ID();
						?>
						<tr><td style="vertical-align: top; padding:10px"><?php the_title(); ?>
							<br>
							<div class="row-actions"><span class="edit"><a href="<?php  echo get_edit_post_link( $post_ID, '' ); ?>" aria-label="Edit “<?php the_title(); ?>”">Edit</a>  | </span><span class="trash"><a href="<?php echo get_delete_post_link( $post_ID, '', false ); ?>" class="submitdelete" aria-label="Move “<?php the_title(); ?>” to the Trash">Trash</a> | </span> </div>
						</td>
						<td style="vertical-align: top; padding:10px"><?php echo get_field('email',$post_ID) ?></td>
						</tr>
						<?php
					
				}
				?>
					</tbody>
				</table>
				<?php
				
				// post pagination
			}
			else {
				// no posts content
			}
			*/
        ?>
    </div>
    <?php
 
}

// Upload alumnis
function alumnis_import_upload(){
    ?>
    <div class="wrap">
		<h2><?php _e( "Alumnis Import" ); ?></h2>
		<p class="alu_import_p">Please upload CSV file with 3 headers ( First name, Last name, Email )</p>
		<form method='get' action="admin.php?page=spee-import-dashboard" id="import-form">
		<!-- enctype="multipart/form-data"   class="dropzone"  -->
		 	<div id="upload_wrapper">
				<div class="dropnupload dropzone" id="import_file">
					<label class="btn btn-blue btn--upload"> Upload</label>
					<div class="img_error_image_import_file"></div>
					<div class="preview_img_import_file"></div>
				</div>
			</div>					
			<div id="importing" style="widht:200px; height:200px; display: none;  left: 0; right: 0; vertical-align:middle;">
			<img src="<?php echo PHP_SPREADSHEET_PLUGIN_URL; ?>/assets/img/preloader.gif" /></div>
			<div id="result"></div>
		</form>
		<script type='text/javascript' src='<?php echo PHP_SPREADSHEET_PLUGIN_URL; ?>/assets/js/dropzone.js'></script>
		<link rel='stylesheet' id='stylea-css'  href='<?php echo PHP_SPREADSHEET_PLUGIN_URL; ?>/assets/css/dropzone.scss' type='text/css' media='all' />
		<script>
			Dropzone.autoDiscover = false;
			jQuery(document).ready(function(){
				// Dropzone.autoDiscover = false;
				var myDropzoneTheSecond = new Dropzone( '#import_file', {
					paramName: 'Filedata',
					url: ajax_url,
					clickable: [ '#import_file', '#import_file *' ],
					enqueueForUpload: false,
					maxFilesize: 500,
					uploadMultiple: false,
					maxFiles: 1,
					addRemoveLinks: true,
					dictDefaultMessage: 'Click anywhere in the box or drop your file here to upload.',
					sending: function( file, xhr, formData ) {
						formData.append( 'action', 'upload_img' );
					},
					accept: function( file, done ) {
						if ( 'text/csv' != file.type && 'application/vnd.ms-excel' != file.type && 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' != file.type  ) {
							jQuery( '.dz-preview' ).html( '' );
							jQuery( '.img_error_image_import_file' ).html( 'Invalid file type.' ).show();
							done( 'Invalid file type.' );
							return false;
						} else {
							jQuery( '.preview_img_import_file' ).html( '' );
							done();
							jQuery( '.img_error_image_import_file' ).html( '' ).hide();
						}
					},
					success: function( data, responseText ) {
						if ( undefined != responseText && '' != responseText ) {
							var data = responseText.split( ':::' );
							jQuery( '.img_error_image_import_file' ).html( '' ).hide();
							jQuery( '#import_file_id' ).val( data[ 0 ] );
							var fileID = data[0];
							//file import
							var importData={'action': 'import_alumni', 'file_id': fileID };
								jQuery.ajax({
								type: 'post',
								// url: ajax_url,
								url: '<?php  echo admin_url('admin-ajax.php'); ?>',
								data: importData,
								beforeSend: function() {
									jQuery("#importing").show();
								},
								success: function( response ) {
										jQuery('#result').html(response);
										jQuery("#importing").hide();
															 
								},
							});
							//end file inport
							jQuery( '.img_error_image_import_file' ).html( '' ).hide();
							jQuery( '#attach_id-error' ).remove();
						}
					},
				});
			})
		</script>
	</div>

<?php
}

// Export data,.
function phpspreadsheet_dashboard() {
	global $wpdb;
	if ( isset( $_GET['export'] )) {
		if ( file_exists(PHP_SPREADSHEET_PLUGIN_DIR . '/vendor/autoload.php') ) {
			// New way.
			$spreadsheet = new Spreadsheet();		
			// Set active sheet index to the first sheet, so Excel opens this as the first sheet
			$spreadsheet->setActiveSheetIndex(0);
			$activeSheet = $spreadsheet->getActiveSheet();
			// Rename worksheet
			if($_GET['user_type']=='no'){	
				$activeSheet->setTitle('National Organizers Excelsheet');			
			} else {
				$activeSheet->setTitle( ucfirst($_GET['user_type']) ??'' . ' Excelsheet ');			
			}
	
			// Add some data as per user type.
			if($_GET['user_type']=='alumni') {

				$a='A';
				$fields=array('ID','Name','Email','Gender','Country','Year','Project','DOB','School','About Me','Facebook','Instagram','Linkedin','Snapchat','Twitter','WhatsApp','Qzone','QQ','WeChat','Means of Transportation','Arriving at', 'Flight Number','Passport Number','Arrival Date & Time','Departure Date & Time','Food restrictions or allergies','Accompanying Adult','Status');
				foreach($fields as $fl) {
					$activeSheet->setCellValue(''.$a.'1', $fl);
					$a++;
				} 
				$activeSheet->getStyle('A1:'.$a.'1')->getFont()->setBold(true);
				$activeSheet->getDefaultColumnDimension('A:'.$a.'')->setWidth(12);	
				// $activeSheet->getColumnDimensionByColumn('A:'.$a.'')->setAutoSize(true);
				$query = "SELECT p.*, u.display_name
					FROM {$wpdb->prefix}posts AS p
					LEFT JOIN {$wpdb->prefix}users AS u ON p.post_author = u.ID
					WHERE p.post_type = 'alumni'  AND (p.post_status = 'publish' OR  p.post_status = 'draft')
					ORDER BY p.ID ASC";				
				$posts   = $wpdb->get_results($query);				
				if ( $posts ) {
					$ia=1;
					foreach ( $posts as $i=>$post ) {
						$facebook=get_field('facebook',$post->ID);
						if(get_field('instagram',$post->ID)!='') {
							$instagram = 'https://www.instagram.com/'.end(explode('/',get_field('instagram',$post->ID)));
						} else {
							$instagram = '';
						}
						// $instagram = get_field('instagram',$post->ID);
						$linkedin = get_field('linkedin',$post->ID);
						$snapchat = get_field('snapchat',$post->ID);
						if(get_field('twitter',$post->ID)!='') {
							$twitter = 'https://twitter.com/'.end(explode('/',get_field('twitter',$post->ID)));
						} else {
							$twitter = '';
						}
						// $twitter = get_field('twitter',$post->ID);
						if(get_field('whatsapp',$post->ID)!='') {
							$whats_app = get_field('whatsapp',$post->ID);
						} else {
							$whats_app = '';
						}
						$qzone = get_field('qzone',$post->ID);
						$qq = get_field('qq',$post->ID);
						$wechat = get_field('wechat',$post->ID);
						$ia++;
						$b='A';
						$questions=get_field('questions',$post->ID);
						$qa='';
						foreach($questions as $question) {
							$qa.=get_the_title($question['question']).'\n\r';
							$qa.=$question['answer'].'\n\r';
						}
						$c1=array('&#8217;','&#8216;');
						$c2=array('&#8220;','&#8221;');
						$travel=get_field('travel_info',$post->ID);
						$val=array(
							$post->ID,
							get_field('first_name',$post->ID).' '.get_field('last_name',$post->ID),
							get_field('email',$post->ID),
							get_field('gender',$post->ID),
							get_field('country',$post->ID),
							get_field('year_competed',$post->ID),
							str_replace('&nbsp;',' ',str_replace('&#8211;','-',str_replace('&amp;','&',str_replace('&#038;','&',str_replace($c1,"'",str_replace($c2,'"',get_the_title(get_field('project_title',$post->ID)))))))),
							get_field('date_of_birth',$post->ID),
							get_field('school',$post->ID),
							str_replace('&nbsp;',' ',str_replace('&#8211;','-',str_replace('&amp;','&',str_replace('&#038;','&',str_replace($c1,"'",str_replace($c2,'"',strip_tags(get_field('about_me',$post->ID)))))))),
							$facebook,$instagram,$linkedin,$snapchat,$twitter,$whats_app,$qzone,$qq,$wechat,$travel[0]['means_of_transportation'],$travel[0]['arriving_at'],$travel[0]['flight_number'],$travel[0]['passport_number'],$travel[0]['arrival_date_&_time'],$travel[0]['departure_date_&_time'],
							get_field('food_allergies',$post->ID), 
							get_field('accompanying_adult',$post->ID), 
							$post->post_status, 
						);
						foreach($val as $v) {
							$activeSheet->setCellValue($b.$ia, $v);
							$b++;
						}
					}
				}



			}
			if($_GET['user_type']=='mentor') {
				$a='A';
				$fields=array('ID','Name','Email','Gender','Address','Expertise','Available','Languages','Looking to Advice','Bio','Status' );
				foreach($fields as $fl) {
					$activeSheet->setCellValue(''.$a.'1', $fl);
					$a++;
				}				
				$activeSheet->getStyle('A1:'.$a.'1')->getFont()->setBold(true);
				$activeSheet->getDefaultColumnDimension('A:'.$a.'')->setWidth(12);
				// $activeSheet->getColumnDimensionByColumn('A:'.$a.'')->setAutoSize(true);
				$query = "SELECT p.*, u.display_name
					FROM {$wpdb->prefix}posts AS p
					LEFT JOIN {$wpdb->prefix}users AS u ON p.post_author = u.ID
					WHERE p.post_type = 'mentor'   AND (p.post_status = 'publish' OR  p.post_status = 'draft')
					ORDER BY p.ID ASC";
				$posts   = $wpdb->get_results($query);
				if ( $posts ) {
					$availability_array = [ 'aa' => 'Available',  'na'=> 'Not Available', 'la' => 'Limited Available' ];
					$language_array = get_mentor_language();
					$ia=1;
					foreach ( $posts as $i=>$post ) {
						$ia++;
						$b='A';
						$experties=get_field('expertise',$post->ID);
						$ex='';
						foreach($experties as $expertie) {
							$ex.=$expertie.', ';
						}
						$languages=get_field('languages',$post->ID);
						$lan='';
						foreach($languages as $language) {
							$trm=get_term_by('id',$language,'mentorlanguage');
							$lan.=$trm->name.', ';
						}
						$val=array(
							$post->ID,
							get_field('first_name',$post->ID).' '.get_field('last_name',$post->ID),
							get_field('email',$post->ID),
							get_field('gender',$post->ID),
							get_field('citystate',$post->ID).' '.get_field('country',$post->ID),
							rtrim($ex,', '),
							$availability_array[get_field('availability',$post->ID)],
							rtrim($lan,', '),
							get_field('looking_to_advice',$post->ID),
							get_field('bio',$post->ID),
							$post->post_status,
						);
						foreach($val as $v) {
							$activeSheet->setCellValue($b.$ia, $v);
							$b++;
						}
					}
				}
			}
			if($_GET['user_type']=='no') {
				$a='A';
				$fields=array('ID','Name','Email','Country','Address','Phone No.','URL','Contact Persons','Sponsors','Users','Facebook','Instagram','Linkedin','Snapchat','Twitter','WhatsApp','Qzone','QQ','WeChat','Status' );
				foreach($fields as $fl) {
					$activeSheet->setCellValue(''.$a.'1', $fl);
					$a++;
				}	
				$activeSheet->getStyle('A1:'.$a.'1')->getFont()->setBold(true);
				$activeSheet->getDefaultColumnDimension('A:'.$a.'')->setWidth(12);
				// $activeSheet->getColumnDimensionByColumn('A:'.$a.'')->setAutoSize(true);
				$query = "SELECT p.*, u.display_name
					FROM {$wpdb->prefix}posts AS p
					LEFT JOIN {$wpdb->prefix}users AS u ON p.post_author = u.ID
					WHERE p.post_type = 'national_organizer'  AND (p.post_status = 'publish' OR  p.post_status = 'draft')
					ORDER BY p.ID ASC";
				$posts   = $wpdb->get_results($query);
				if ( $posts ) {
					$availability_array = [ 'aa' => 'Available',  'na'=> 'Not Available', 'la' => 'Limited Available' ];
					$language_array = get_mentor_language();
					$ia=1;
					foreach ( $posts as $i=>$post ) {
						$facebook=get_field('facebook',$post->ID);
						if(get_field('instagram',$post->ID)!='') {
							$instagram = 'https://www.instagram.com/'.end(explode('/',get_field('instagram',$post->ID)));
						} else {
							$instagram = '';
						}
						// $instagram = get_field('instagram',$post->ID);
						$linkedin = get_field('linkedin',$post->ID);
						$snapchat = get_field('snapchat',$post->ID);
						if(get_field('twitter',$post->ID)!='') {
							$twitter = 'https://twitter.com/'.end(explode('/',get_field('twitter',$post->ID)));
						} else {
							$twitter = '';
						}
						// $twitter = get_field('twitter',$post->ID);
						if(get_field('whatsapp',$post->ID)!='') {
							$whats_app = get_field('whatsapp',$post->ID);
						} else {
							$whats_app = '';
						}
						$qzone = get_field('qzone',$post->ID);
						$qq = get_field('qq',$post->ID);
						$wechat = get_field('wechat',$post->ID);
						$ia++;
						$b='A';
						$contact_person=get_field('contact_persons',$post->ID);
						$ex='';
						foreach($contact_person as $person) {
							$ex.=$person['name'].'('.$person['email_address'].'), ';
						}
						$no_users=get_field('user_lists',$post->ID);
						$nusers='';
						foreach($no_users as $usr) {							 
							$nusers.=$usr['email_address'].', ';
						}
						$sponsors=get_field('sponsors',$post->ID);
						$spon='';
						foreach($sponsors as $spn) {
							$spon.=$spn['sponsor'].'(URL:'.$spn['url'].'), ';
						}
						$addressnew=get_field('postal_address',$post->ID);
						$val=array(
								$post->ID,
								get_field('name_of_no',$post->ID),
								get_field('email_address',$post->ID),
								get_field('country',$post->ID),
								$addressnew['address'],
								get_field('phone_no',$post->ID),
								get_field('url',$post->ID),
								rtrim($ex,', '),
								rtrim($spon,', '),
								rtrim($nusers,', '),
								$facebook,$instagram,$linkedin,$snapchat,$twitter,$whats_app,$qzone,$qq,$wechat,
								$post->post_status,
							);
						foreach($val as $v) {
							$activeSheet->setCellValue($b.$ia, $v);
							$b++;
						}					 
					}
				}				
			}
			if($_GET['user_type']=='project') {
				$a='A';
				$fields=array('ID','Project Name','Water issue addressed','Country','Year Completed','Project Owner','Authors','Contact Persons','A1 Poster','A3 Poster','20 page scientific paper','Status');
				foreach($fields as $fl) {
					$activeSheet->setCellValue(''.$a.'1', $fl);
					$a++;
				}
				$activeSheet->getStyle('A1:'.$a.'1')->getFont()->setBold(true);
				$activeSheet->getDefaultColumnDimension('A:'.$a.'')->setWidth(12);
				// $activeSheet->getColumnDimensionByColumn('A:'.$a.'')->setAutoSize(true);
				$query = "SELECT p.*, u.display_name
					FROM {$wpdb->prefix}posts AS p
					LEFT JOIN {$wpdb->prefix}users AS u ON p.post_author = u.ID
					WHERE p.post_type = 'projects'  AND (p.post_status = 'publish' OR  p.post_status = 'draft')
					ORDER BY p.ID ASC";
				$posts   = $wpdb->get_results($query);
				echo count( $posts);
				if ( $posts ) {
					$availability_array = [ 'aa' => 'Available',  'na'=> 'Not Available', 'la' => 'Limited Available' ];
					$language_array = get_mentor_language();
					$ia=1;
					foreach ( $posts as $i=>$post ) {
						$ia++;
						$b='A';
						$authors=get_field('author',$post->ID);
						$ex='';
						foreach($authors as $person) {
							$ex.=get_field('first_name',$person).' '.get_field('last_name',$person).', ';
						}
						$water_issue=get_field('water_issue_addressed',$post->ID);
						$trm= get_term_by( 'id', $water_issue, 'waterissue' );
						$watteris=$trm->name;
						$a1_poster='No';
						$a3_poster='No';
						$a20_poster='No';
						if(get_field('a1_poster_type',$post->ID)=='link' && get_field('a1_poster',$post->ID)!='') {
								$a1_poster='Yes';
						}
						if(get_field('a1_poster_type',$post->ID)=='file' && get_field('a1_poster_file',$post->ID)!='') {
								$a1_poster='Yes';
						}
						if(get_field('a3_poster_type',$post->ID)=='link' && get_field('a3_poster_link',$post->ID)!='') {
							$a3_poster='Yes';
						}
						if(get_field('a3_poster_type',$post->ID)=='file' && get_field('a3_poster',$post->ID)!='') {
							$a3_poster='Yes';
						}
						if(get_field('20_page_scientific_paper_type',$post->ID)=='link' && get_field('20_page_scientific_paper_link',$post->ID)!='') {
							$a20_poster='Yes';
						}
						if(get_field('20_page_scientific_paper_type',$post->ID)=='file' && get_field('20_page_scientific_paper',$post->ID)!='') {
							$a20_poster='Yes';
						}
						$c1=array('&#8217;','&#8216;');
						$c2=array('&#8220;','&#8221;');
						$val=array(
							$post->ID,
							str_replace('&nbsp;',' ',str_replace('&#8211;','-',str_replace('&amp;','&',str_replace('&#038;','&',str_replace($c1,"'",str_replace($c2,'"',get_the_title($post->ID))))))),
							$watteris,
							get_field('country',$post->ID),
							get_field('year_competed',$post->ID),
							get_field('first_name',get_field('project_owner',$post->ID)).' '.get_field('last_name',get_field('project_owner',$post->ID)),
							rtrim($ex,', '),
							get_field('first_name',get_field('contact_person',$post->ID)).' '.get_field('last_name',get_field('contact_person',$post->ID)),
							$a1_poster,
							$a3_poster,
							$a20_poster,
							$post->post_status,
						);
						foreach($val as $v) {
							$activeSheet->setCellValue($b.$ia, $v);
							$b++;
						}
					}
				}
			}
			if ( $_GET['user_type'] == 'voting_list' ) {
				$a = 'A';
				$fields = array( 'ID','Project Name','Email', 'IP', 'Date' );
				foreach($fields as $fl) {
					$activeSheet->setCellValue(''.$a.'1', $fl);
					$a++;
				}
				$activeSheet->getStyle('A1:'.$a.'1')->getFont()->setBold(true);
				$activeSheet->getDefaultColumnDimension('A:'.$a.'')->setWidth(12);
				// $activeSheet->getColumnDimensionByColumn('A:'.$a.'')->setAutoSize(true);

				$where = ' WHERE 1=1 ';
				$current_year     = date( 'Y' );
				$project_selected = isset( $_GET['project'] ) && ! empty( $_GET['project'] ) ? $_GET['project'] : '';
				$email_selected   = isset( $_GET['current_voter_email'] ) && ! empty( $_GET['current_voter_email'] ) ? $_GET['current_voter_email'] : '';
				$ip_selected      = isset( $_GET['current_voter_ip'] ) && ! empty( $_GET['current_voter_ip'] ) ? $_GET['current_voter_ip'] : '';
				if( '' != trim( $project_selected )) {
					$project_selected = explode( '&ppid=', $project_selected );
						$project_selected = $project_selected[1];
					// $psql    = "SELECT ID from ".$wpdb->prefix."posts join " . $wpdb->prefix. "postmeta on ".$wpdb->prefix."posts.ID = ".$wpdb->prefix."postmeta.post_id where post_title like('%".$project_selected."%') AND post_type = 'projects' and meta_key = 'year_competed' and meta_value = " . $current_year;
					$psql    = "SELECT ID from ".$wpdb->prefix."posts join " . $wpdb->prefix. "postmeta on ".$wpdb->prefix."posts.ID = ".$wpdb->prefix."postmeta.post_id where post_type = 'projects' and meta_key = 'year_competed' and meta_value = " . $current_year . " and ID = " . $project_selected;
					$presult = $wpdb->get_results( $psql, 'ARRAY_A' );
					$p_array = [];
					foreach($presult as $p ) {
						$p_array[] = $p['ID'];
					}
					$p_string = implode( ',', $p_array );
					$where    .= ' AND project_id IN( '.$p_string.' )';
				} else {
					$query = "SELECT ID from ".$wpdb->prefix."posts join " . $wpdb->prefix. "postmeta on ".$wpdb->prefix."posts.ID = ".$wpdb->prefix."postmeta.post_id where post_type = 'projects' and meta_key = 'year_competed' and meta_value = " . $current_year;
					$res = $wpdb->get_results( $query, 'ARRAY_A' );
					$array = [];
					foreach( $res as $value ) {
						$array[] = $value['ID'];
					}
					$string = implode( ',', $array );
					$where .= ' AND project_id IN('.$string.')';
				}
				if( '' !== trim( $email_selected ) ) {
						$where .= " AND email LIKE( '%" . $email_selected . "%' )";
				}
				if( '' !== trim( $ip_selected ) ) {
					$where .= " AND ip LIKE( '%" . $ip_selected ."%' )";
				}
				$sql = "SELECT * FROM ".$wpdb->prefix."voting ".$where."  ORDER BY date DESC";
				$posts = $wpdb->get_results( $sql );
				if ( isset( $posts ) && ! empty( $posts ) ) {
					$ia = 1;
					foreach( $posts as $i => $post ) {
						$ia++;
						$b = 'A';
						$val = array(
							$post->ID,
							get_the_title( $post->project_id ),
							$post->email,
							$post->ip,
							$post->date
						);
						foreach($val as $v) {
							$activeSheet->setCellValue($b.$ia, $v);
							$b++;
						}
					}
					
				}
			}
			
			// Write a new .xlsx file
			// Redirect output to a client’s web browser
			ob_clean();
			ob_start();
			switch ( $_GET['format'] ) { // Write a new file based on format.
				case 'csv':			
					// Redirect output to a client’s web browser (CSV)
					$writer = IOFactory::createWriter($spreadsheet, 'Csv');		
					header('Content-type: text/csv; charset=ISO-8859-2');
					// header('Content-type: text/csv');
					header("Cache-Control: no-store, no-cache");
					header('Content-Encoding: ISO-8859-2');
					echo "\xEF\xBB\xBF"; // UTF-8 BOM
					$filename = $_GET['user_type'].'_exported_excel_' . time() . '.csv';
					if($_GET['user_type']=='no'){		
						$filename = 'national_organizers_excelsheet_exported_excel_' . time() . '.csv';		
					} else {		
						$filename = $_GET['user_type'].'_exported_excel_' . time() . '.csv';	
					}
					$headerContent = 'Content-Disposition: attachment;filename="' . $filename . '"';
					header($headerContent);
					$writer->save('php://output');					
					break;
				case 'xls':
					// Redirect output to a client’s web browser (Excel5)
					$writer = IOFactory::createWriter($spreadsheet, 'Xls');			
					header('Content-Type: application/vnd.ms-excel');
					// header('Content-Type: text/xls');
					header("Cache-Control: no-store, no-cache");
					if($_GET['user_type']=='no'){		
						$filename = 'national_organizers_excelsheet_exported_excel_' . time() . '.xls';		
					} else {		
						$filename = $_GET['user_type'].'_exported_excel_' . time() . '.xls';	
					}
					$headerContent = 'Content-Disposition: attachment;filename="' . $filename . '"';
					header($headerContent);
					$writer->save('php://output');					
					break;
				case 'xlsx':
					// Redirect output to a client’s web browser (Excel2007)
					$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');			
					header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
					// header('Content-Type: text/xlsx');
					header("Cache-Control: no-store, no-cache");
					if($_GET['user_type']=='no'){		
						$filename = 'national_organizers_excelsheet_exported_excel_' . time() . '.xlsx';		
					} else {		
						$filename = $_GET['user_type'].'_exported_excel_' . time() . '.xlsx';	
					}
					$headerContent = 'Content-Disposition: attachment;filename="' . $filename . '"';
					header($headerContent);
					$writer->save('php://output');					
					break;
			}
			exit;
		}
	}


	
	$test = array(
        array(
            'key' => 'testdata',
            'value' => '',
            'compare' => 'NOT EXISTS',
        ),
        array(
        'key' => 'year_competed',
        'value' => date( 'Y' ),
        'compare' => '=',
        )
    );
	$projects_lists_auto = new WP_Query([
	    'posts_per_page' => -1,
	    'offset' => 0,
	    'post_type' => 'projects',
	    'post_status'    => 'publish',
	    'orderby' => 'meta_value post_date',
	    'meta_key' => 'year_competed', 
	    'order' => 'DESC',
	    'meta_query' => $test,
	]);
?>
<div class="wrap">
	<h2><?php _e( "PhpSpreadsheet Export" ); ?></h2>
	<form method='get' action="admin.php?page=phpspreadsheet-dashboard" id="export-form">
		<input type="hidden" name='page' value="phpspreadsheet-dashboard"/>
		<input type="hidden" name='noheader' value="1"/>
		<label>Export Type</label><br>
		<select name="user_type" required="required" id="user_type">
			<option value="0">--Select User Type--</option>
			<option value="alumni">Alumnis</option>
			<option value="mentor">Mentors</option>
			<option value="no">National organizers</option>
			<option value="project">Projects</option>
			<option value="voting_list">Vote</option>
		</select><br>
		<br><div style="display: none;" id="voting_list_options">
			<label>Voting Filter</label><br>
			Project: <input type="text" value="" name="project"  id="projecttext" />
			Email: <input type="text" value="" name="current_voter_email" />
			IP: <input type="text" value="" name="current_voter_ip" />
		</div><br>
		<div style="">
			<label>Export to</label><br><br>
			<label for"formatCSV" class="export-type" ><input type="radio" name='format' id="formatCSV"  value="csv" checked="checked"/>  csv</label>
			<label for"formatXLS" class="export-type" ><input type="radio" name='format' id="formatXLS"  value="xls"/>  xls</label>
			<label for"formatXLSX" class="export-type" ><input type="radio" name='format' id="formatXLSX" value="xlsx"/>  xslx</label>
		</div><br>
		<input type="submit" name='export' id="csvExport" value="Export"/>
	</form>
	<style> .export-type:first{ padding-top: 10px; } .export-type { font-size: 15px;   text-transform: capitalize; padding-right: 10px; } </style>
	<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
	<script>
		jQuery( function() {
			var availableTags = [
				<?php
				foreach($projects_lists_auto->posts as $post) {
					// echo '"'.preg_replace('/[^\p{L}\p{N}\s-]/u', '',$post->post_title).'",';
					echo '"'.preg_replace('/[^\p{L}\p{N}\s:?~-]/u', '',$post->post_title).'&ppid='.$post->ID.'",';
				}
				?>
			];
			jQuery( '#projecttext' ).autocomplete({
				source: availableTags
			}).focus(function() {
				jQuery(this).autocomplete('search', " ")
			});
		});
 		jQuery(document).ready(function(){
 			jQuery( '#user_type' ).change( function() {
 				var usr_type = jQuery('#user_type').val();
 				if ( 'voting_list' == usr_type ) {
					jQuery( '#voting_list_options' ).css( "display", "block" );
				} else {
					jQuery( '#voting_list_options' ).css( "display", "none" );
				}
 			})
			jQuery('#export-form').submit(function(){
				var usr_type = jQuery('#user_type').val();
				if (usr_type == '0' ) {
			  	alert("Please select export type");
					return false;
				}
			});
		});
	 </script>
</div>
<?php
}




// For ajax communication inside this file.
add_action('wp_ajax_import_alumni', 'import_alumni' );
add_action('wp_ajax_nopriv_import_alumni', 'import_alumni' );
// Check if Spreadsheet class exists.
if(class_exists( Spreadsheet::class ) ) {	
	function getSheets($fileName) {
		try {			
			$fileType = IOFactory::identify($fileName);			
			$reader = IOFactory::createReader($fileType);
			// if($fileType == 'Csv' || $fileType == 'csv') {
			// 	// $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
			// 	$reader = IOFactory::createReader('Csv');
			// } else if($fileType == 'Xlsx' || $fileType == 'xlsx') {
			// 	// $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
			// 	$reader = IOFactory::createReader('Xlsx');
			// } else if($fileType == 'Xls' || $fileType == 'xls') {
			// 	// $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
			// 	$reader = IOFactory::createReader('Xls');
			// } else {
			// 	$reader = IOFactory::createReader($fileType);
			// }
			$reader->setReadDataOnly(true);

			$spreadsheet = $reader->load($fileName);
			// $all = $spreadsheet->getSheetNames()[1];
			$sheets = [];
			foreach ($spreadsheet->getAllSheets() as $key => $sheet) {
				$sheets[$sheet->getTitle()] = $sheet->toArray();
			}
			return $sheets;
		} catch (Exception $e) {
			 die($e->getMessage());
		}
	}
	
} 

function import_alumni() {
	
	$alumni_file=get_attached_file( $_POST['file_id'] );
	$alumnies = getSheets($alumni_file);
	$i=0;
	$output = '<ul>';
	if(!empty($alumnies['Worksheet'] ) && is_array($alumnies['Worksheet'] ) && strtolower($alumnies['Worksheet'][0][0]) == 'first_name' && strtolower($alumnies['Worksheet'][0][1]) == 'last_name' && strtolower($alumnies['Worksheet'][0][2]) == 'email' ) {	

		foreach($alumnies['Worksheet'] as $alumni)		{  
			
			if( $i >= 1 && filter_var($alumni[2], FILTER_VALIDATE_EMAIL) ){
				$user= ' <span class="user--name"> ' .$alumni[0].' '.$alumni[1].' </span><span class="user--email"> '.$alumni[2].' </span> ';
				$post_count = check_email_existence( $alumni[2] );

				if ( $post_count >= 1 )
				{
					$status= ' <span style="color:#ff0000;">Email Already Exist !</span>';
				
				} else {
					$args=array(
						'post_title' => $alumni[0].' '.$alumni[1],
						'post_type' => 'alumni',
					);
					$id=wp_insert_post($args);
					update_field('first_name',$alumni[0],$id);
					update_field('last_name',$alumni[1],$id);
					update_field('email',$alumni[2],$id);
					$key=md5(uniqid(rand(), true));
					add_post_meta($id,'key',$key);
					add_post_meta($id,'status','0');
					new_alumni_notification($alumni[0],$alumni[1],$alumni[2],$id,$key);
					if($id)
					{
						$status= ' <span style="color:#09a509;">Success!</span>';
					}
					else
					{
						$status= ' <span style="color:#ff0000;">Sorry Unable to import alumni !</span>';
					}
				}
				
				
			}
			$i++;

			$output.= '<li><b>'.$user.'</b>-'.$status.'</li>';
			
		}
	} else {
		$output.= '<li style="color:#ff0000;"><br>Data formatting is not matching, please fromat it and try again!</li>';
	}
	$output.=  '</ul>';
	echo $output;
	// exit();
	wp_die();
		
}
