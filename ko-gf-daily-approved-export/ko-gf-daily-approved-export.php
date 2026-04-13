<?php
/**
 * Plugin Name: KO – GF Daily Approved Export
 * Description: Sends a daily XLSX export of approved Gravity Forms entries that have not yet been exported.
 * Version: 1.0.8
 * Author: KO
 * License: GPL-2.0-or-later
 * Text Domain: ko-gfde
 */

if ( ! defined( 'ABSPATH' ) ) exit;

class KO_GF_Daily_Approved_Export {

	const OPTION_SETTINGS = 'ko_gfde_settings';
	const OPTION_STATUS   = 'ko_gfde_status';
	const CRON_HOOK       = 'ko_gfde_run_scheduled_export';
	const NONCE_ACTION    = 'ko_gfde_admin_action';

	// Per-entry flags stored with GF entry meta.
	const META_EXPORTED_AT    = 'ko_gfde_exported_at';
	const META_EXPORTED_BATCH = 'ko_gfde_exported_batch';
	const META_EXPORTED_FILE  = 'ko_gfde_exported_file';

	/**
	 * Export columns.
	 */
	public static function get_export_fields() {
		return [
			['id' => '20',                  'label' => 'Customer Address (Full)'],
			['id' => 'id',                  'label' => 'Entry Id'],
			['id' => '96',                  'label' => 'Make'],
			['id' => '47',                  'label' => 'Model'],
			['id' => '60',                  'label' => 'Year'],
			['id' => '81',                  'label' => 'Mileage'],
			['id' => '1',                   'label' => 'VIN'],
			['id' => 'workflow_final_status','label' => 'Status'],
			['id' => '105',                 'label' => 'Created On'],
			['id' => '68',                  'label' => 'Dealer Code'],
			['id' => '69',                  'label' => 'Dealer Name'],
			['id' => '3.3',                 'label' => 'Salesperson Name'],
			['id' => '17',                  'label' => 'Salesperson Email'],
			['id' => '93',                  'label' => 'Customer Name'],
			['id' => '18',                  'label' => 'Customer Email'],
			['id' => '13',                  'label' => 'PO Number'],
			['id' => '104',                 'label' => 'Odometer Start'],
			['id' => '11',                  'label' => 'Engine'],
			['id' => '12',                  'label' => 'Transmission'],
			['id' => '103',                 'label' => 'Resale Value'],
			['id' => '101',                 'label' => 'Delivery Date'],
			['id' => '99',                  'label' => 'Plan'],
			['id' => '111',                 'label' => 'Coverage Code'],
		];
	}

	public static function bootstrap() {
		$self = new self();

		register_activation_hook( __FILE__, [ __CLASS__, 'activate' ] );
		register_deactivation_hook( __FILE__, [ __CLASS__, 'deactivate' ] );

		add_action( 'admin_menu', [ $self, 'admin_menu' ] );
		add_action( 'admin_init', [ $self, 'handle_post' ] );

		add_action( self::CRON_HOOK, [ $self, 'run_scheduled_export' ] );
	}

	public static function activate() {
		$settings = self::get_settings();
		self::schedule_event( $settings );
	}

	public static function deactivate() {
		wp_clear_scheduled_hook( self::CRON_HOOK );
	}

	public static function get_settings() {
		$defaults = [
			'form_id'       => 2,
			'send_to'       => get_option( 'admin_email', '' ),
			'send_cc'       => '',
			'schedule'      => '01:30',
			'filename'      => 'Daily-Warranty-Approval-Report-{date}',
			'email_message' => "Attached is the daily approved export for {form_title}.\n\nEntries Included: {count}\nGenerated: {generated}",
			'active'        => 1,
		];

		$settings = get_option( self::OPTION_SETTINGS, [] );
		if ( ! is_array( $settings ) ) {
			$settings = [];
		}

		return wp_parse_args( $settings, $defaults );
	}

	public static function update_status( $data ) {
		$status = get_option( self::OPTION_STATUS, [] );
		if ( ! is_array( $status ) ) {
			$status = [];
		}
		$status = array_merge( $status, $data );
		update_option( self::OPTION_STATUS, $status, false );
	}

	public static function get_status() {
		$status = get_option( self::OPTION_STATUS, [] );
		return is_array( $status ) ? $status : [];
	}

	public static function schedule_event( $settings = null ) {
		if ( ! $settings ) {
			$settings = self::get_settings();
		}

		wp_clear_scheduled_hook( self::CRON_HOOK );

		if ( empty( $settings['active'] ) ) {
			return;
		}

		$time_str = ! empty( $settings['schedule'] ) ? $settings['schedule'] : '01:30';
		$parts = explode( ':', $time_str );
		$hour  = isset( $parts[0] ) ? max( 0, min( 23, intval( $parts[0] ) ) ) : 1;
		$min   = isset( $parts[1] ) ? max( 0, min( 59, intval( $parts[1] ) ) ) : 30;

		$tz   = wp_timezone();
		$now  = new DateTime( 'now', $tz );
		$next = new DateTime( 'now', $tz );
		$next->setTime( $hour, $min, 0 );

		if ( $next <= $now ) {
			$next->modify( '+1 day' );
		}

		wp_schedule_event( $next->getTimestamp(), 'daily', self::CRON_HOOK );
	}

	public function admin_menu() {
		add_management_page(
			'KO Daily Approved Export',
			'KO Daily Approved Export',
			'manage_options',
			'ko-gfde',
			[ $this, 'render_admin_page' ]
		);
	}

	public function handle_post() {
		if ( ! is_admin() || ! current_user_can( 'manage_options' ) ) {
			return;
		}

		if ( empty( $_POST['ko_gfde_action'] ) ) {
			return;
		}

		check_admin_referer( self::NONCE_ACTION );

		$action = sanitize_text_field( wp_unslash( $_POST['ko_gfde_action'] ) );
		$redirect = admin_url( 'tools.php?page=ko-gfde' );

		if ( $action === 'save_settings' ) {
			$settings = self::get_settings();
			$settings['form_id']  = isset( $_POST['form_id'] ) ? absint( $_POST['form_id'] ) : 0;
			$settings['send_to']  = isset( $_POST['send_to'] ) ? sanitize_text_field( wp_unslash( $_POST['send_to'] ) ) : '';
			$settings['send_cc']  = isset( $_POST['send_cc'] ) ? sanitize_text_field( wp_unslash( $_POST['send_cc'] ) ) : '';
			$settings['schedule'] = isset( $_POST['schedule'] ) ? sanitize_text_field( wp_unslash( $_POST['schedule'] ) ) : '01:30';
			$settings['filename']      = isset( $_POST['filename'] ) ? sanitize_text_field( wp_unslash( $_POST['filename'] ) ) : 'Daily-Warranty-Approval-Report-{date}';
			$settings['email_message'] = isset( $_POST['email_message'] ) ? wp_kses_post( wp_unslash( $_POST['email_message'] ) ) : '';
			$settings['active']        = ! empty( $_POST['active'] ) ? 1 : 0;

			update_option( self::OPTION_SETTINGS, $settings, false );
			self::schedule_event( $settings );

			wp_safe_redirect( add_query_arg( 'ko_gfde_notice', 'saved', $redirect ) );
			exit;
		}

		if ( $action === 'run_now' ) {
			$result = $this->run_export( true );
			$code = ! empty( $result['success'] ) ? 'ran' : 'run_error';
			set_transient( 'ko_gfde_last_manual_result', $result, 300 );
			wp_safe_redirect( add_query_arg( 'ko_gfde_notice', $code, $redirect ) );
			exit;
		}

		if ( $action === 'clear_specific_flags' ) {
			$raw_ids = isset( $_POST['entry_ids'] ) ? sanitize_text_field( wp_unslash( $_POST['entry_ids'] ) ) : '';
			$ids = array_filter( array_unique( array_map( 'absint', preg_split( '/\s*,\s*/', $raw_ids ) ) ) );
			$cleared = 0;
			foreach ( $ids as $entry_id ) {
				if ( $this->clear_export_flag_for_entry( $entry_id ) ) {
					$cleared++;
				}
			}
			wp_safe_redirect( add_query_arg( [ 'ko_gfde_notice' => 'specific_reset', 'ko_gfde_reset_count' => $cleared ], $redirect ) );
			exit;
		}

		if ( $action === 'reset_flags' ) {
			$count = $this->reset_export_flags();
			wp_safe_redirect( add_query_arg( [ 'ko_gfde_notice' => 'reset', 'ko_gfde_reset_count' => $count ], $redirect ) );
			exit;
		}

		if ( $action === 'clear_single_flag' ) {
			$entry_id = isset( $_POST['entry_id'] ) ? absint( $_POST['entry_id'] ) : 0;
			if ( $entry_id ) {
				$this->clear_export_flag_for_entry( $entry_id );
			}
			wp_safe_redirect( add_query_arg( 'ko_gfde_notice', 'single_reset', $redirect ) );
			exit;
		}
	}

	public function render_admin_page() {
		$settings = self::get_settings();
		$status   = self::get_status();
		$manual   = get_transient( 'ko_gfde_last_manual_result' );
		$forms    = class_exists( 'GFAPI' ) ? GFAPI::get_forms() : [];
		?>
		<div class="wrap ko-gfde-admin">
			<style>
				.ko-gfde-admin{max-width:1200px}
				.ko-gfde-grid{display:grid;grid-template-columns:2fr 1fr;gap:20px;align-items:start;margin-top:18px}
				.ko-gfde-card{background:#fff;border:1px solid #dcdcde;border-radius:14px;padding:22px;box-shadow:0 1px 2px rgba(16,24,40,.04)}
				.ko-gfde-card h2{margin-top:0;margin-bottom:14px}
				.ko-gfde-muted{color:#50575e}
				.ko-gfde-kpis{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
				.ko-gfde-kpi{background:#f6f7f7;border-radius:12px;padding:14px}
				.ko-gfde-kpi span{display:block;color:#50575e;font-size:12px;text-transform:uppercase;letter-spacing:.04em}
				.ko-gfde-kpi strong{display:block;font-size:18px;margin-top:4px;line-height:1.3}
				.ko-gfde-actions form{display:inline-block;margin-right:10px}
				.ko-gfde-admin textarea{width:100%;min-height:130px}
				.ko-gfde-admin input.regular-text,.ko-gfde-admin select,.ko-gfde-admin input[type="time"]{min-width:320px}
				.ko-gfde-help{font-size:12px;color:#646970;margin-top:6px}
				@media (max-width: 960px){.ko-gfde-grid{grid-template-columns:1fr}.ko-gfde-kpis{grid-template-columns:1fr}}
			</style>

			<h1>KO Daily Approved Export</h1>
			<p class="ko-gfde-muted">Custom XLSX exporter for approved Gravity Forms entries that have not yet been exported by this plugin. Runs independently of Entry Automation.</p>

			<?php if ( ! empty( $_GET['ko_gfde_notice'] ) ) : ?>
				<?php $notice = sanitize_text_field( wp_unslash( $_GET['ko_gfde_notice'] ) ); ?>
				<div class="notice notice-success is-dismissible">
					<p>
						<?php
						if ( $notice === 'saved' ) {
							echo esc_html( 'Settings saved.' );
						} elseif ( $notice === 'ran' ) {
							echo esc_html( 'Manual export run completed.' );
						} elseif ( $notice === 'run_error' ) {
							echo esc_html( 'Manual export run completed with an issue. See result below.' );
						} elseif ( $notice === 'reset' ) {
							$count = isset( $_GET['ko_gfde_reset_count'] ) ? absint( $_GET['ko_gfde_reset_count'] ) : 0;
							echo esc_html( sprintf( 'Export flags reset for %d entries.', $count ) );
						} elseif ( $notice === 'single_reset' ) {
							echo esc_html( 'Export flag cleared for the selected entry.' );
						} elseif ( $notice === 'specific_reset' ) {
							$count = isset( $_GET['ko_gfde_reset_count'] ) ? absint( $_GET['ko_gfde_reset_count'] ) : 0;
							echo esc_html( sprintf( 'Export flag cleared for %d entries.', $count ) );
						}
						?>
					</p>
				</div>
			<?php endif; ?>

			<div class="ko-gfde-grid">
				<div class="ko-gfde-card">
					<h2>Settings</h2>
					<form method="post">
						<?php wp_nonce_field( self::NONCE_ACTION ); ?>
						<input type="hidden" name="ko_gfde_action" value="save_settings" />
						<table class="form-table" role="presentation">
							<tr>
								<th scope="row"><label for="form_id">Form</label></th>
								<td>
									<select name="form_id" id="form_id">
										<?php foreach ( $forms as $form ) : ?>
											<option value="<?php echo esc_attr( $form['id'] ); ?>" <?php selected( (int) $settings['form_id'], (int) $form['id'] ); ?>>
												<?php echo esc_html( '#' . $form['id'] . ' — ' . rgar( $form, 'title' ) ); ?>
											</option>
										<?php endforeach; ?>
									</select>
									<div class="ko-gfde-help">Current site uses one daily export configuration. This keeps the form selectable for future flexibility.</div>
								</td>
							</tr>
							<tr>
								<th scope="row"><label for="send_to">To</label></th>
								<td>
									<input type="text" class="regular-text" name="send_to" id="send_to" value="<?php echo esc_attr( $settings['send_to'] ); ?>" />
									<div class="ko-gfde-help">Comma-separated email addresses.</div>
								</td>
							</tr>
							<tr>
								<th scope="row"><label for="send_cc">CC</label></th>
								<td>
									<input type="text" class="regular-text" name="send_cc" id="send_cc" value="<?php echo esc_attr( $settings['send_cc'] ); ?>" />
									<div class="ko-gfde-help">Optional comma-separated email addresses.</div>
								</td>
							</tr>
							<tr>
								<th scope="row"><label for="schedule">Schedule Time</label></th>
								<td>
									<input type="time" name="schedule" id="schedule" value="<?php echo esc_attr( $settings['schedule'] ); ?>" />
									<div class="ko-gfde-help">Daily local site time. Requested default is 1:30 AM.</div>
								</td>
							</tr>
							<tr>
								<th scope="row"><label for="filename">File Name Pattern</label></th>
								<td>
									<input type="text" class="regular-text" name="filename" id="filename" value="<?php echo esc_attr( $settings['filename'] ); ?>" />
									<div class="ko-gfde-help">Use <code>{date}</code> to append the run date. Example: <code>Daily-Warranty-Approval-Report-{date}</code></div>
								</td>
							</tr>
							<tr>
								<th scope="row"><label for="email_message">Email Message</label></th>
								<td>
									<textarea name="email_message" id="email_message"><?php echo esc_textarea( $settings['email_message'] ); ?></textarea>
									<div class="ko-gfde-help">Supported tokens: <code>{form_title}</code>, <code>{count}</code>, <code>{generated}</code>, <code>{site_name}</code>, <code>{file_name}</code></div>
								</td>
							</tr>
							<tr>
								<th scope="row">Active</th>
								<td>
									<label><input type="checkbox" name="active" value="1" <?php checked( ! empty( $settings['active'] ) ); ?> /> Enable scheduled export</label>
								</td>
							</tr>
						</table>
						<?php submit_button( 'Save Settings' ); ?>
					</form>
				</div>

				<div>
					<div class="ko-gfde-card">
						<h2>Run Status</h2>
						<div class="ko-gfde-kpis">
							<div class="ko-gfde-kpi"><span>Last Run</span><strong><?php echo ! empty( $status['last_run'] ) ? esc_html( $status['last_run'] ) : '—'; ?></strong></div>
							<div class="ko-gfde-kpi"><span>Last Result</span><strong><?php echo ! empty( $status['last_result'] ) ? esc_html( $status['last_result'] ) : '—'; ?></strong></div>
							<div class="ko-gfde-kpi"><span>Last Export Count</span><strong><?php echo isset( $status['last_count'] ) ? esc_html( $status['last_count'] ) : '—'; ?></strong></div>
							<div class="ko-gfde-kpi"><span>Next Scheduled Run</span><strong><?php echo esc_html( $this->get_next_run_display() ); ?></strong></div>
						</div>
						<p style="margin-top:16px;"><strong>Last File:</strong> <?php echo ! empty( $status['last_file'] ) ? esc_html( $status['last_file'] ) : '—'; ?></p>

						<div class="ko-gfde-actions" style="margin-top:18px;">
							<form method="post">
								<?php wp_nonce_field( self::NONCE_ACTION ); ?>
								<input type="hidden" name="ko_gfde_action" value="run_now" />
								<?php submit_button( 'Run Export Now', 'secondary', 'submit', false ); ?>
							</form>
						</div>
					</div>

					<div class="ko-gfde-card" style="margin-top:20px;">
						<h2>Reset Entry Flags</h2>
						<p class="ko-gfde-muted">Enter one or more Gravity Forms Entry IDs separated by commas. Those entries will be re-queued for the next scheduled export only.</p>
						<form method="post">
							<?php wp_nonce_field( self::NONCE_ACTION ); ?>
							<input type="hidden" name="ko_gfde_action" value="clear_specific_flags" />
							<input type="text" class="regular-text" name="entry_ids" placeholder="473 or 473,474,481" />
							<?php submit_button( 'Reset Entry Flags', 'secondary', 'submit', false ); ?>
						</form>

						<form method="post" style="margin-top:14px;" onsubmit="return confirm('Reset export flags for all entries on the selected form?');">
							<?php wp_nonce_field( self::NONCE_ACTION ); ?>
							<input type="hidden" name="ko_gfde_action" value="reset_flags" />
							<?php submit_button( 'Reset All Export Flags', 'delete', 'submit', false ); ?>
						</form>
					</div>
				</div>
			</div>

			<?php if ( is_array( $manual ) ) : ?>
				<div class="ko-gfde-card" style="margin-top:20px;">
					<h2>Last Manual Run Result</h2>
					<pre style="background:#fff;border:1px solid #ccd0d4;padding:12px;max-width:1000px;overflow:auto;"><?php echo esc_html( print_r( $manual, true ) ); ?></pre>
				</div>
			<?php endif; ?>
		</div>
		<?php
	}
	public function get_next_run_display() {
		$ts = wp_next_scheduled( self::CRON_HOOK );
		if ( ! $ts ) {
			return 'Not scheduled';
		}
		return wp_date( 'F j, Y \a\t g:i A T', $ts, wp_timezone() );
	}

	public function run_scheduled_export() {
		$this->run_export( false );
	}

	public function run_export( $manual = false ) {
		$settings = self::get_settings();
		$form_id  = absint( $settings['form_id'] );

		$result = [
			'success' => false,
			'manual'  => $manual,
			'form_id' => $form_id,
			'count'   => 0,
			'file'    => '',
			'message' => '',
			'entries' => [],
		];

		if ( ! class_exists( 'GFAPI' ) || ! function_exists( 'gform_get_meta' ) || ! function_exists( 'gform_update_meta' ) ) {
			$result['message'] = 'Gravity Forms API/meta functions are not available.';
			$this->log_status( $result );
			return $result;
		}

		if ( empty( $settings['send_to'] ) ) {
			$result['message'] = 'No recipient email address is configured.';
			$this->log_status( $result );
			return $result;
		}

		$entry_ids = $this->get_approved_unexported_entry_ids( $form_id );
		$result['entries'] = $entry_ids;
		$result['count']   = count( $entry_ids );

		if ( empty( $entry_ids ) ) {
			$result['success'] = true;
			$result['message'] = 'No approved unexported entries found.';
			$this->log_status( $result );
			return $result;
		}

		$headers = array_map(
			function( $field ) {
				return $field['label'];
			},
			self::get_export_fields()
		);

		$rows = [];
		$entries = [];
		foreach ( $entry_ids as $entry_id ) {
			$entry = GFAPI::get_entry( $entry_id );
			if ( is_wp_error( $entry ) || empty( $entry ) ) {
				continue;
			}
			$entries[] = $entry;
			$rows[] = $this->build_export_row( $entry );
		}

		if ( empty( $rows ) ) {
			$result['success'] = true;
			$result['message'] = 'No exportable rows were built.';
			$this->log_status( $result );
			return $result;
		}

		$file_path = $this->generate_xlsx_file( $headers, $rows, $settings );
		if ( is_wp_error( $file_path ) ) {
			$result['message'] = $file_path->get_error_message();
			$this->log_status( $result );
			return $result;
		}

		$result['file'] = basename( $file_path );

		$subject = 'Daily Warranty Approval Report - ' . wp_date( 'Y-m-d', time(), wp_timezone() );
		$form   = GFAPI::get_form( $form_id );
		$tokens = [
			'{form_title}' => is_array( $form ) ? (string) rgar( $form, 'title' ) : (string) $form_id,
			'{count}'      => (string) count( $rows ),
			'{generated}'  => wp_date( 'F j, Y \a\t g:i A T', time(), wp_timezone() ),
			'{site_name}'  => wp_specialchars_decode( get_bloginfo( 'name' ), ENT_QUOTES ),
			'{file_name}'  => basename( $file_path ),
		];
		$body_template = ! empty( $settings['email_message'] ) ? (string) $settings['email_message'] : 'Attached is the daily approved export for {form_title}.\n\nEntries Included: {count}\nGenerated: {generated}';
		$body = strtr( $body_template, $tokens );

		$to = $this->parse_emails( $settings['send_to'] );
		$headers_mail = [];
		$cc = $this->parse_emails( $settings['send_cc'] );
		if ( ! empty( $cc ) ) {
			$headers_mail[] = 'Cc: ' . implode( ',', $cc );
		}

		$sent = wp_mail( $to, $subject, $body, $headers_mail, [ $file_path ] );

		if ( ! $sent ) {
			$result['message'] = 'wp_mail returned false. Entries were not marked exported.';
			$this->log_status( $result );
			return $result;
		}

		$batch = wp_date( 'Y-m-d', time(), wp_timezone() );
		foreach ( $entries as $entry ) {
			gform_update_meta( $entry['id'], self::META_EXPORTED_AT, gmdate( 'Y-m-d H:i:s' ) );
			gform_update_meta( $entry['id'], self::META_EXPORTED_BATCH, $batch );
			gform_update_meta( $entry['id'], self::META_EXPORTED_FILE, basename( $file_path ) );
		}

		$result['success'] = true;
		$result['message'] = 'Export sent successfully and entries marked exported.';
		$this->log_status( $result );

		return $result;
	}

	private function parse_emails( $value ) {
		$parts = array_filter( array_map( 'trim', explode( ',', (string) $value ) ) );
		return array_values(
			array_filter(
				$parts,
				function( $email ) {
					return is_email( $email );
				}
			)
		);
	}

	private function log_status( $result ) {
		$data = [
			'last_run'    => wp_date( 'F j, Y \a\t g:i A T', time(), wp_timezone() ),
			'last_result' => ! empty( $result['message'] ) ? $result['message'] : ( ! empty( $result['success'] ) ? 'Success' : 'Failed' ),
			'last_count'  => isset( $result['count'] ) ? intval( $result['count'] ) : 0,
			'last_file'   => ! empty( $result['file'] ) ? $result['file'] : '',
		];

		self::update_status( $data );

		if ( defined( 'WP_DEBUG_LOG' ) && WP_DEBUG_LOG ) {
			error_log( '[KO GFDE] ' . wp_json_encode( $result ) );
		}
	}

	private function get_approved_unexported_entry_ids( $form_id ) {
		global $wpdb;

		$entry_table = $this->get_gf_entry_table();
		$meta_table  = $this->get_gf_entry_meta_table();

		// Build candidate entry IDs from workflow approval state.
		$sql = $wpdb->prepare(
			"SELECT DISTINCT e.id
			 FROM {$entry_table} e
			 INNER JOIN {$meta_table} m ON m.entry_id = e.id
			 WHERE e.form_id = %d
			   AND e.status = 'active'
			   AND (
					(m.meta_key = 'workflow_final_status' AND LOWER(m.meta_value) IN ('complete','approved'))
					OR
					(m.meta_key = 'workflow_step_status_4' AND LOWER(m.meta_value) = 'approved')
			   )
			 ORDER BY e.id ASC",
			$form_id
		);

		$ids = $wpdb->get_col( $sql );
		if ( empty( $ids ) ) {
			return [];
		}

		$ids = array_map( 'intval', $ids );

		$filtered = [];
		foreach ( $ids as $entry_id ) {
			$exported_at = gform_get_meta( $entry_id, self::META_EXPORTED_AT );
			if ( empty( $exported_at ) ) {
				$filtered[] = $entry_id;
			}
		}

		return $filtered;
	}

	private function clear_export_flag_for_entry( $entry_id ) {
		$entry_id = absint( $entry_id );
		if ( ! $entry_id ) {
			return false;
		}

		gform_delete_meta( $entry_id, self::META_EXPORTED_AT );
		gform_delete_meta( $entry_id, self::META_EXPORTED_BATCH );
		gform_delete_meta( $entry_id, self::META_EXPORTED_FILE );

		return true;
	}

	private function reset_export_flags() {
		$form_id = absint( self::get_settings()['form_id'] );
		$ids = $this->get_all_form_entry_ids( $form_id );
		$count = 0;

		foreach ( $ids as $entry_id ) {
			if ( gform_get_meta( $entry_id, self::META_EXPORTED_AT ) !== '' && gform_get_meta( $entry_id, self::META_EXPORTED_AT ) !== null ) {
				$count++;
			}
			gform_delete_meta( $entry_id, self::META_EXPORTED_AT );
			gform_delete_meta( $entry_id, self::META_EXPORTED_BATCH );
			gform_delete_meta( $entry_id, self::META_EXPORTED_FILE );
		}

		return $count;
	}

	private function get_all_form_entry_ids( $form_id ) {
		global $wpdb;
		$entry_table = $this->get_gf_entry_table();

		$sql = $wpdb->prepare(
			"SELECT id FROM {$entry_table} WHERE form_id = %d AND status = 'active' ORDER BY id ASC",
			$form_id
		);

		$ids = $wpdb->get_col( $sql );
		return array_map( 'intval', $ids );
	}

	private function build_export_row( $entry ) {
		$row = [];
		foreach ( self::get_export_fields() as $field ) {
			$row[] = $this->get_entry_export_value( $entry, $field['id'] );
		}
		return $row;
	}

	private function get_entry_export_value( $entry, $field_id ) {
		switch ( $field_id ) {
			case 'id':
				return (int) rgar( $entry, 'id' );

			case 'workflow_final_status':
				return strtolower( (string) gform_get_meta( $entry['id'], 'workflow_final_status' ) );

			case '20':
				return $this->build_full_address( $entry, '20' );

			default:
				$value = rgar( $entry, (string) $field_id );

				// Fall back to entry meta for workflow-ish fields or derived fields.
				if ( $value === '' || $value === null ) {
					$meta = gform_get_meta( $entry['id'], (string) $field_id );
					if ( $meta !== '' && $meta !== null ) {
						$value = $meta;
					}
				}

				// Normalize arrays / objects.
				if ( is_array( $value ) ) {
					$value = implode( ', ', array_map( 'strval', $value ) );
				} elseif ( is_object( $value ) ) {
					$value = wp_json_encode( $value );
				}

				return $value;
		}
	}


	private function build_full_address( $entry, $base_field_id ) {
		$parts = [
			trim( (string) rgar( $entry, $base_field_id . '.1' ) ),
			trim( (string) rgar( $entry, $base_field_id . '.2' ) ),
		];
		$street = implode( ', ', array_filter( $parts ) );

		$city    = trim( (string) rgar( $entry, $base_field_id . '.3' ) );
		$state   = trim( (string) rgar( $entry, $base_field_id . '.4' ) );
		$zip     = trim( (string) rgar( $entry, $base_field_id . '.5' ) );
		$country = trim( (string) rgar( $entry, $base_field_id . '.6' ) );

		$line2_parts = [];
		if ( $city !== '' ) {
			$line2_parts[] = $city;
		}
		$state_zip = trim( implode( ' ', array_filter( [ $state, $zip ] ) ) );
		if ( $state_zip !== '' ) {
			$line2_parts[] = $state_zip;
		}
		$line2 = implode( ', ', $line2_parts );

		$all = array_filter( [ $street, $line2, $country ] );
		return implode( ', ', $all );
	}

	private function generate_xlsx_file( $headers, $rows, $settings ) {
		$uploads = wp_upload_dir();
		if ( ! empty( $uploads['error'] ) ) {
			return new WP_Error( 'ko_gfde_uploads', $uploads['error'] );
		}

		$dir = trailingslashit( $uploads['basedir'] ) . 'ko-gfde';
		if ( ! wp_mkdir_p( $dir ) ) {
			return new WP_Error( 'ko_gfde_dir', 'Could not create export directory.' );
		}

		$date = wp_date( 'Y-m-d', time(), wp_timezone() );
		$base = ! empty( $settings['filename'] ) ? $settings['filename'] : 'Daily-Warranty-Approval-Report-{date}';
		$base = str_replace( '{date}', $date, $base );
		$base = sanitize_file_name( $base );

		if ( substr( $base, -5 ) !== '.xlsx' ) {
			$base .= '.xlsx';
		}

		$path = trailingslashit( $dir ) . $base;

		$writer = new KO_GFDE_XLSX_Writer();
		$ok = $writer->write( $path, 'Worksheet', $headers, $rows );

		if ( ! $ok ) {
			return new WP_Error( 'ko_gfde_xlsx', 'Could not generate XLSX file.' );
		}

		return $path;
	}

	private function get_gf_entry_table() {
		global $wpdb;
		$new_table    = $wpdb->prefix . 'gf_entry';
		$legacy_table = $wpdb->prefix . 'rg_lead';

		$exists = $wpdb->get_var( $wpdb->prepare( "SHOW TABLES LIKE %s", $new_table ) );
		if ( $exists ) return $new_table;

		$exists = $wpdb->get_var( $wpdb->prepare( "SHOW TABLES LIKE %s", $legacy_table ) );
		if ( $exists ) return $legacy_table;

		return $new_table;
	}

	private function get_gf_entry_meta_table() {
		global $wpdb;
		$new_table = $wpdb->prefix . 'gf_entry_meta';
		$exists = $wpdb->get_var( $wpdb->prepare( "SHOW TABLES LIKE %s", $new_table ) );
		if ( $exists ) return $new_table;

		// Modern installs should have gf_entry_meta. Fallback for safety.
		return $new_table;
	}
}

/**
 * Minimal XLSX writer using ZipArchive and inline strings.
 * No external library required.
 */
class KO_GFDE_XLSX_Writer {

	public function write( $filepath, $sheet_name, $headers, $rows ) {
		if ( ! class_exists( 'ZipArchive' ) ) {
			return false;
		}

		$zip = new ZipArchive();
		if ( $zip->open( $filepath, ZipArchive::CREATE | ZipArchive::OVERWRITE ) !== true ) {
			return false;
		}

		$sheet_xml = $this->build_sheet_xml( $headers, $rows );

		$zip->addFromString( '[Content_Types].xml', $this->content_types_xml() );
		$zip->addEmptyDir( '_rels' );
		$zip->addFromString( '_rels/.rels', $this->rels_xml() );

		$zip->addEmptyDir( 'docProps' );
		$zip->addFromString( 'docProps/app.xml', $this->app_xml( $sheet_name ) );
		$zip->addFromString( 'docProps/core.xml', $this->core_xml() );

		$zip->addEmptyDir( 'xl' );
		$zip->addFromString( 'xl/workbook.xml', $this->workbook_xml( $sheet_name ) );
		$zip->addFromString( 'xl/styles.xml', $this->styles_xml() );
		$zip->addEmptyDir( 'xl/_rels' );
		$zip->addFromString( 'xl/_rels/workbook.xml.rels', $this->workbook_rels_xml() );

		$zip->addEmptyDir( 'xl/worksheets' );
		$zip->addFromString( 'xl/worksheets/sheet1.xml', $sheet_xml );

		$zip->close();

		return file_exists( $filepath );
	}

	private function build_sheet_xml( $headers, $rows ) {
		$xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
		$xml .= '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
		$xml .= '<sheetData>';

		$all_rows = array_merge( [ $headers ], $rows );
		$row_num = 1;
		foreach ( $all_rows as $row ) {
			$xml .= '<row r="' . $row_num . '">';
			$col_num = 1;
			foreach ( $row as $cell ) {
				$cell_ref = $this->col_name( $col_num ) . $row_num;
				if ( is_numeric( $cell ) && (string)(int)$cell === (string)$cell || is_float( $cell ) || (is_string($cell) && preg_match('/^-?\d+(\.\d+)?$/', $cell)) ) {
					$xml .= '<c r="' . esc_attr( $cell_ref ) . '" t="n"><v>' . $this->xml( (string) $cell ) . '</v></c>';
				} else {
					$xml .= '<c r="' . esc_attr( $cell_ref ) . '" t="inlineStr"><is><t>' . $this->xml( (string) $cell ) . '</t></is></c>';
				}
				$col_num++;
			}
			$xml .= '</row>';
			$row_num++;
		}

		$xml .= '</sheetData></worksheet>';

		return $xml;
	}

	private function content_types_xml() {
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
			. '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
			. '<Default Extension="xml" ContentType="application/xml"/>'
			. '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
			. '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
			. '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
			. '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
			. '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
			. '</Types>';
	}

	private function rels_xml() {
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
			. '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
			. '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
			. '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
			. '</Relationships>';
	}

	private function app_xml( $sheet_name ) {
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
			. '<Application>ChatGPT</Application>'
			. '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>'
			. '<TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>' . $this->xml( $sheet_name ) . '</vt:lpstr></vt:vector></TitlesOfParts>'
			. '</Properties>';
	}

	private function core_xml() {
		$now = gmdate( 'Y-m-d\TH:i:s\Z' );
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
			. '<dc:creator>KO GF Daily Approved Export</dc:creator>'
			. '<cp:lastModifiedBy>KO GF Daily Approved Export</cp:lastModifiedBy>'
			. '<dcterms:created xsi:type="dcterms:W3CDTF">' . $now . '</dcterms:created>'
			. '<dcterms:modified xsi:type="dcterms:W3CDTF">' . $now . '</dcterms:modified>'
			. '</cp:coreProperties>';
	}

	private function workbook_xml( $sheet_name ) {
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
			. '<sheets><sheet name="' . $this->xml( $sheet_name ) . '" sheetId="1" r:id="rId1"/></sheets>'
			. '</workbook>';
	}

	private function workbook_rels_xml() {
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
			. '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
			. '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
			. '</Relationships>';
	}

	private function styles_xml() {
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			. '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
			. '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
			. '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
			. '<borders count="1"><border/></borders>'
			. '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
			. '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
			. '</styleSheet>';
	}

	private function col_name( $num ) {
		$name = '';
		while ( $num > 0 ) {
			$num--;
			$name = chr( 65 + ( $num % 26 ) ) . $name;
			$num = intval( $num / 26 );
		}
		return $name;
	}

	private function xml( $value ) {
		return htmlspecialchars( $value, ENT_XML1 | ENT_COMPAT, 'UTF-8' );
	}
}

KO_GF_Daily_Approved_Export::bootstrap();
