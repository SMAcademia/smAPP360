<?php
/**
 * Plugin Name: SM Academia — Inscripciones a Google Sheets
 * Description: Envía las inscripciones y preinscripciones de los formularios Jetpack al GAS de SM Academia.
 * Version:     1.2.0
 * Author:      SM Academia
 *
 * INSTALACIÓN:
 *   1. Subir este archivo a /wp-content/plugins/sm-academia-inscripciones/
 *      (crear la carpeta si no existe)
 *   2. Activar el plugin en WP Admin → Plugins
 *
 * REQUISITO:
 *   El GAS debe tener la acción NUEVA_INSCRIPCION en doPost.
 *   Ver instrucciones en el chat de SM Academia.
 */

if ( ! defined( 'ABSPATH' ) ) exit;

// ── URL del GAS ────────────────────────────────────────────────────────────────
define( 'SM_GAS_URL', 'https://script.google.com/macros/s/AKfycbyhFDN-zhSurQwdxHomYm8mIkdjDiFqbWYKoYSkkw5GH6nZ9vhrS5fZfvLTqYgEwUoROA/exec' );

// ── Email del administrador para notificaciones de error ───────────────────────
define( 'SM_ADMIN_EMAIL', 'asmarriv1986@gmail.com' );

/**
 * Se dispara cuando Jetpack Contact Form envía un formulario.
 *
 * @param int   $post_id    ID de la página que contiene el formulario.
 * @param array $mail       Configuración del correo del formulario.
 * @param array $all_values Array asociativo [ 'Etiqueta del campo' => 'valor' ].
 */
add_action( 'grunion_after_message_sent', 'sm_formulario_a_gas', 10, 3 );
function sm_formulario_a_gas( $post_id, $mail, $all_values ) {
    $slug = get_post_field( 'post_name', $post_id );

    if ( $slug === 'inscripcion-academia-sm-futbol' ) {
        sm_inscripcion_a_gas( $post_id, $all_values );
    } elseif ( $slug === 'preinscripcion-escaleritas' ) {
        sm_preinscripcion_a_gas( $post_id, $all_values );
    } elseif ( (int) $post_id === 7182 || $slug === 'inscripcion-ceip-atlantida' ) {
        sm_preinscripcion_atlantida_a_gas( $post_id, $all_values );
    }
}

// ── Inscripción SM Academia Fútbol ──────────────────────────────────────────
function sm_inscripcion_a_gas( $post_id, $all_values ) {

    // Días preferidos puede venir como array (checkboxes) o string
    $dias_raw = $all_values['Días preferidos'] ?? '';
    if ( is_array( $dias_raw ) ) {
        $dias = implode( ', ', array_map( 'sanitize_text_field', $dias_raw ) );
    } else {
        $dias = sanitize_text_field( $dias_raw );
    }

    $data = [
        'action'             => 'NUEVA_INSCRIPCION',
        'nombre_alumno'      => sanitize_text_field( $all_values['Nombre del jugador/a']       ?? '' ),
        'apellidos_alumno'   => sanitize_text_field( $all_values['Apellidos del jugador/a']    ?? '' ),
        'fecha_nac'          => sanitize_text_field( $all_values['Fecha nacimiento (AAAA-MM-DD)'] ?? '' ),
        'colegio'            => sanitize_text_field( $all_values['Colegio']                    ?? '' ),
        'observaciones'      => sanitize_textarea_field( $all_values['Observaciones']          ?? '' ),
        'nombre_tutor'       => sanitize_text_field( $all_values['Nombre del tutor/a']         ?? '' ),
        'apellidos_tutor'    => sanitize_text_field( $all_values['Apellidos del tutor/a']      ?? '' ),
        'telefono'           => sanitize_text_field( $all_values['Teléfono']                   ?? '' ),
        'email'              => sanitize_email(      $all_values['Correo electrónico']         ?? '' ),
        'servicio'           => sanitize_text_field( $all_values['Servicio']                   ?? '' ),
        'modalidad'          => sanitize_text_field( $all_values['Modalidad']                  ?? '' ),
        'preferencia_centro' => sanitize_text_field( $all_values['Preferencias']               ?? '' ),
        'dias'               => $dias,
        'autoriza_imagen'    => sanitize_text_field( $all_values['Autorizo uso de imagen']     ?? 'Sí' ),
        'mensaje'            => sanitize_textarea_field( $all_values['Mensaje']                ?? '' ),
    ];

    // Llamada al GAS — no bloqueante: el usuario no espera respuesta
    $response = wp_remote_post( SM_GAS_URL, [
        'headers'  => [ 'Content-Type' => 'application/json; charset=utf-8' ],
        'body'     => wp_json_encode( $data ),
        'timeout'  => 30,
        'blocking' => false,   // respuesta ignorada; fallo silencioso en frontend
    ] );

    // Log de error (solo visible en debug.log si WP_DEBUG_LOG está activo)
    if ( is_wp_error( $response ) ) {
        error_log( '[SM Academia] Error al llamar al GAS (inscripción): ' . $response->get_error_message() );
    }
}

// ── Preinscripción de interés — SM Extraescolares Escaleritas ───────────────
function sm_preinscripcion_a_gas( $post_id, $all_values ) {

    // Actividades puede llegar como array (checkboxes múltiples)
    $actividades_raw = $all_values['Actividades de interés'] ?? '';
    if ( is_array( $actividades_raw ) ) {
        $actividades = implode( ', ', array_map( 'sanitize_text_field', $actividades_raw ) );
    } else {
        $actividades = sanitize_text_field( $actividades_raw );
    }

    $dias_raw = $all_values['Días disponibles'] ?? '';
    if ( is_array( $dias_raw ) ) {
        $dias = implode( ', ', array_map( 'sanitize_text_field', $dias_raw ) );
    } else {
        $dias = sanitize_text_field( $dias_raw );
    }

    $data = [
        'action'              => 'NUEVA_PREINSCRIPCION',
        'centro'              => 'ESCALERITAS',
        'curso_academico'     => '2026-2027',
        // Alumno
        'nombre_alumno'       => sanitize_text_field( $all_values['Nombre del alumno/a']         ?? '' ),
        'apellidos_alumno'    => sanitize_text_field( $all_values['Apellidos del alumno/a']      ?? '' ),
        'curso_alumno'        => sanitize_text_field( $all_values['Curso escolar']               ?? '' ),
        'fecha_nac'           => sanitize_text_field( $all_values['Fecha de nacimiento']         ?? '' ),
        'salud'               => sanitize_textarea_field( $all_values['Enfermedades/alergias']   ?? '' ),
        'observaciones'       => sanitize_textarea_field( $all_values['Observaciones']           ?? '' ),
        // Tutor 1
        'nombre_tutor1'       => sanitize_text_field( $all_values['Nombre tutor 1']              ?? '' ),
        'apellidos_tutor1'    => sanitize_text_field( $all_values['Apellidos tutor 1']           ?? '' ),
        'relacion_tutor1'     => sanitize_text_field( $all_values['Relación tutor 1']            ?? '' ),
        'nif_tutor1'          => sanitize_text_field( $all_values['NIF/DNI tutor 1']             ?? '' ),
        'telefono_tutor1'     => sanitize_text_field( $all_values['Teléfono tutor 1']            ?? '' ),
        'email_tutor1'        => sanitize_email(      $all_values['Email tutor 1']               ?? '' ),
        // Tutor 2
        'nombre_tutor2'       => sanitize_text_field( $all_values['Nombre tutor 2']              ?? '' ),
        'apellidos_tutor2'    => sanitize_text_field( $all_values['Apellidos tutor 2']           ?? '' ),
        'relacion_tutor2'     => sanitize_text_field( $all_values['Relación tutor 2']            ?? '' ),
        'nif_tutor2'          => sanitize_text_field( $all_values['NIF/DNI tutor 2']             ?? '' ),
        'telefono_tutor2'     => sanitize_text_field( $all_values['Teléfono tutor 2']            ?? '' ),
        'email_tutor2'        => sanitize_email(      $all_values['Email tutor 2']               ?? '' ),
        // Consentimientos
        'autoriza_imagen'     => sanitize_text_field( $all_values['Cesión de imágenes']          ?? '' ),
        'acepta_datos'        => sanitize_text_field( $all_values['Política de datos']           ?? '' ),
        'acepta_datos_salud'  => sanitize_text_field( $all_values['Consentimiento datos salud']  ?? '' ),
        // AMPA y pago
        'socio_ampa'          => sanitize_text_field( $all_values['¿Desea hacerse socio del AMPA?'] ?? '' ),
        'iban'                => sanitize_text_field( $all_values['IBAN para domiciliación']     ?? '' ),
        // Actividades
        'actividades'         => $actividades,
        'dias_disponibles'    => $dias,
    ];

    $response = wp_remote_post( SM_GAS_URL, [
        'headers'  => [ 'Content-Type' => 'application/json; charset=utf-8' ],
        'body'     => wp_json_encode( $data ),
        'timeout'  => 30,
        'blocking' => false,
    ] );

    if ( is_wp_error( $response ) ) {
        error_log( '[SM Academia] Error al llamar al GAS (preinscripción): ' . $response->get_error_message() );
    }
}

// ── Preinscripción CEIP Atlántida (página 7182) ─────────────────────────────
function sm_preinscripcion_atlantida_a_gas( $post_id, $all_values ) {

    // Actividades de interés (checkboxes múltiples)
    $acti_raw = $all_values['Actividades de interés'] ?? '';
    if ( is_array( $acti_raw ) ) {
        $actividades = implode( ', ', array_map( 'sanitize_text_field', $acti_raw ) );
    } else {
        $actividades = sanitize_text_field( $acti_raw );
    }

    // NEAE: se concatena en el campo salud para no requerir columna nueva en Sheets
    $salud_base = sanitize_textarea_field( $all_values['Enfermedades/alergias'] ?? '' );
    $neae_val   = sanitize_text_field( $all_values['El alumno/a tiene NEAE (necesidades específicas de apoyo educativo), como TEA, TDAH, dislexia, altas capacidades, discapacidad u otras'] ?? '' );
    $neae_aclar = sanitize_textarea_field( $all_values['Aclaraciones sobre las NEAE (opcional)'] ?? '' );
    $salud = $salud_base;
    if ( $neae_val ) {
        $salud .= ( $salud ? "\n" : '' ) . 'NEAE: ' . $neae_val;
        if ( $neae_aclar ) {
            $salud .= "\nAclaraciones NEAE: " . $neae_aclar;
        }
    }

    // Acepta datos salud: el label puede incluir "y NEAE" en la página actualizada
    $acepta_salud = sanitize_text_field(
        $all_values['Consentimiento datos salud y NEAE']
        ?? $all_values['Consentimiento datos salud']
        ?? ''
    );

    // AMPA Adargoma: probamos la etiqueta específica y la genérica como fallback
    $socio_ampa = sanitize_text_field(
        $all_values['¿Desea hacerse socio del AMPA Adargoma?']
        ?? $all_values['¿Desea hacerse socio del AMPA?']
        ?? ''
    );

    $data = [
        'action'              => 'NUEVA_PREINSCRIPCION',
        'centro'              => 'CEIP ATLÁNTIDA',
        'curso_academico'     => '2026-2027',
        // Alumno
        'nombre_alumno'       => sanitize_text_field(     $all_values['Nombre del alumno/a']       ?? '' ),
        'apellidos_alumno'    => sanitize_text_field(     $all_values['Apellidos del alumno/a']    ?? '' ),
        'curso_alumno'        => sanitize_text_field(     $all_values['Curso escolar']             ?? '' ),
        'fecha_nac'           => sanitize_text_field(     $all_values['Fecha de nacimiento']       ?? '' ),
        'salud'               => $salud,
        'observaciones'       => sanitize_textarea_field( $all_values['Observaciones']             ?? '' ),
        // Tutor 1
        'nombre_tutor1'       => sanitize_text_field(     $all_values['Nombre tutor 1']            ?? '' ),
        'apellidos_tutor1'    => sanitize_text_field(     $all_values['Apellidos tutor 1']         ?? '' ),
        'relacion_tutor1'     => sanitize_text_field(     $all_values['Relación tutor 1']          ?? '' ),
        'nif_tutor1'          => sanitize_text_field(     $all_values['NIF/DNI tutor 1']           ?? '' ),
        'telefono_tutor1'     => sanitize_text_field(     $all_values['Teléfono tutor 1']          ?? '' ),
        'email_tutor1'        => sanitize_email(          $all_values['Email tutor 1']             ?? '' ),
        // Tutor 2
        'nombre_tutor2'       => sanitize_text_field(     $all_values['Nombre tutor 2']            ?? '' ),
        'apellidos_tutor2'    => sanitize_text_field(     $all_values['Apellidos tutor 2']         ?? '' ),
        'relacion_tutor2'     => sanitize_text_field(     $all_values['Relación tutor 2']          ?? '' ),
        'nif_tutor2'          => sanitize_text_field(     $all_values['NIF/DNI tutor 2']           ?? '' ),
        'telefono_tutor2'     => sanitize_text_field(     $all_values['Teléfono tutor 2']          ?? '' ),
        'email_tutor2'        => sanitize_email(          $all_values['Email tutor 2']             ?? '' ),
        // Consentimientos
        'autoriza_imagen'     => sanitize_text_field(     $all_values['Cesión de imágenes']        ?? '' ),
        'acepta_datos'        => sanitize_text_field(     $all_values['Política de datos']         ?? '' ),
        'acepta_datos_salud'  => $acepta_salud,
        // AMPA y pago
        'socio_ampa'          => $socio_ampa,
        'iban'                => sanitize_text_field(     $all_values['IBAN para domiciliación']   ?? '' ),
        // Actividades
        'actividades'         => $actividades,
    ];

    $response = wp_remote_post( SM_GAS_URL, [
        'headers'  => [ 'Content-Type' => 'application/json; charset=utf-8' ],
        'body'     => wp_json_encode( $data ),
        'timeout'  => 30,
        'blocking' => false,
    ] );

    if ( is_wp_error( $response ) ) {
        error_log( '[SM Academia] Error al llamar al GAS (preinscripción Atlántida): ' . $response->get_error_message() );
    }

    // Acuse a la familia — el fallo no impide el envío al GAS (ya fue no-blocking)
    sm_acuse_atlantida( $all_values );
}

// ── Acuse de recibo — CEIP Atlántida ────────────────────────────────────────
function sm_acuse_atlantida( $all_values ) {
    $nombre_alumno    = sanitize_text_field( $all_values['Nombre del alumno/a']    ?? '' );
    $apellidos_alumno = sanitize_text_field( $all_values['Apellidos del alumno/a'] ?? '' );
    $nombre_completo  = trim( $nombre_alumno . ' ' . $apellidos_alumno );
    $nombre_tutor1    = sanitize_text_field( $all_values['Nombre tutor 1']         ?? '' );
    $email_tutor1     = sanitize_email(      $all_values['Email tutor 1']          ?? '' );

    $acti_raw = $all_values['Actividades de interés'] ?? '';
    if ( is_array( $acti_raw ) ) {
        $items         = array_map( 'sanitize_text_field', $acti_raw );
        $servicios_html = '<ul style="margin:4px 0;padding-left:20px;">'
            . implode( '', array_map( function( $s ) {
                return '<li style="color:#333;">' . esc_html( $s ) . '</li>';
            }, $items ) )
            . '</ul>';
    } else {
        $servicios_html = '<p style="margin:0;color:#333;">' . esc_html( sanitize_text_field( $acti_raw ) ) . '</p>';
    }

    if ( ! $email_tutor1 ) {
        error_log( '[SM Academia] Acuse Atlántida: sin email tutor 1 para alumno ' . $nombre_completo );
        return;
    }

    $subject = 'Hemos recibido la inscripción de ' . $nombre_completo . ' — CEIP Atlántida';
    $headers = [
        'Content-Type: text/html; charset=UTF-8',
        'Cc: direccion@sm-academia.com',
    ];
    $body = sm_acuse_atlantida_html(
        esc_html( $nombre_completo ),
        esc_html( $nombre_tutor1 ),
        $servicios_html
    );

    try {
        wp_mail( $email_tutor1, $subject, $body, $headers );
    } catch ( Exception $e ) {
        error_log( '[SM Academia] wp_mail error (acuse Atlántida): ' . $e->getMessage() );
    }
}

// ── HTML del acuse — CEIP Atlántida ─────────────────────────────────────────
function sm_acuse_atlantida_html( $nombre_alumno, $nombre_tutor1, $servicios_html ) {
    return '<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
</head>
<body style="margin:0;padding:0;background:#f4f6fb;font-family:Arial,Helvetica,sans-serif;">
<div style="max-width:600px;margin:0 auto;">

  <div style="background:#1a3a6b;padding:28px 32px;text-align:center;">
    <p style="margin:0;color:#ffffff;font-size:22px;font-weight:bold;">SM Academia</p>
    <p style="margin:6px 0 0;color:#a8c4f0;font-size:13px;">
      CEIP Atlántida &middot; Actividades Extraescolares 2026&ndash;2027
    </p>
  </div>

  <div style="background:#ffffff;padding:32px;">
    <p style="font-size:16px;color:#222;">Hola ' . $nombre_tutor1 . ',</p>

    <p style="color:#444;line-height:1.6;">
      Hemos recibido correctamente la inscripción de
      <strong>' . $nombre_alumno . '</strong>
      en las actividades extraescolares del <strong>CEIP Atlántida</strong>
      para el curso 2026&ndash;2027,
      coordinadas por el <strong>AMPA Adargoma</strong>.
    </p>

    <div style="background:#eef3fb;border-left:4px solid #1a3a6b;padding:14px 18px;
                margin:20px 0;border-radius:0 6px 6px 0;">
      <p style="margin:0 0 8px;font-weight:bold;color:#1a3a6b;">Actividades solicitadas:</p>
      ' . $servicios_html . '
    </div>

    <h2 style="color:#1a3a6b;font-size:16px;border-bottom:1px solid #d0d8ef;
               padding-bottom:6px;margin-top:28px;">
      Cómo realizar el pago
    </h2>
    <p style="color:#444;line-height:1.6;">
      El pago se realiza <strong>mensualmente, entre los días&nbsp;1 y&nbsp;10 de cada mes</strong>.
      Puede usar cualquiera de estas opciones:
    </p>

    <table style="width:100%;border-collapse:collapse;margin:12px 0;font-size:14px;">
      <tr style="background:#f7f9ff;">
        <td style="padding:12px 14px;border:1px solid #d0d8ef;
                   font-weight:bold;color:#1a3a6b;width:35%;vertical-align:top;">
          📱 Bizum
        </td>
        <td style="padding:12px 14px;border:1px solid #d0d8ef;color:#333;">
          <strong style="font-size:18px;letter-spacing:1px;">625&nbsp;46&nbsp;82&nbsp;96</strong>
          <br><span style="font-size:12px;color:#666;">Concepto: nombre del alumno/a</span>
        </td>
      </tr>
      <tr>
        <td style="padding:12px 14px;border:1px solid #d0d8ef;
                   font-weight:bold;color:#1a3a6b;vertical-align:top;">
          🏦 Transferencia bancaria
        </td>
        <td style="padding:12px 14px;border:1px solid #d0d8ef;color:#333;">
          <span style="font-family:monospace;font-size:14px;letter-spacing:0.5px;">
            ES69 0182 5036 7102 0163 5388
          </span>
          <br><span style="font-size:12px;color:#666;">Concepto: nombre del alumno/a</span>
        </td>
      </tr>
    </table>

    <h2 style="color:#1a3a6b;font-size:16px;border-bottom:1px solid #d0d8ef;
               padding-bottom:6px;margin-top:28px;">
      Horario del servicio de recogida tard&iacute;a
    </h2>
    <p style="color:#444;">De <strong>15:00 a 17:00&nbsp;h</strong>.</p>

    <h2 style="color:#1a3a6b;font-size:16px;border-bottom:1px solid #d0d8ef;
               padding-bottom:6px;margin-top:28px;">
      &iquest;Tiene alguna pregunta?
    </h2>
    <p style="color:#444;line-height:1.6;">
      Escr&iacute;banos a
      <a href="mailto:direccion@sm-academia.com" style="color:#1a3a6b;">
        direccion@sm-academia.com
      </a>
      y le responderemos a la mayor brevedad.
    </p>

    <p style="margin-top:28px;color:#444;">
      Un cordial saludo,<br>
      <strong>SM Academia</strong>
    </p>
  </div>

  <div style="background:#f0f2f8;padding:14px;text-align:center;
              font-size:11px;color:#888;line-height:1.6;">
    SM Academia &middot; Las Palmas de Gran Canaria &middot;
    <a href="https://sm-academia.com" style="color:#888;">sm-academia.com</a><br>
    Este correo se envi&oacute; autom&aacute;ticamente al registrar la inscripci&oacute;n.
  </div>

</div>
</body>
</html>';
}
