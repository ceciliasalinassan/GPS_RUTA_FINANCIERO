import nodemailer from "nodemailer";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ ok: false, error: "Método no permitido" });
  }

  try {
    const { to, subject, body, attachments = [] } = req.body || {};

    if (!to || !subject || !body) {
      return res.status(400).json({ ok: false, error: "Faltan datos: destinatario, asunto o mensaje." });
    }

    const user = process.env.OUTLOOK_USER || "gpsruta007@outlook.com";
    const pass = process.env.OUTLOOK_APP_PASSWORD;

    if (!pass) {
      return res.status(500).json({
        ok: false,
        error: "Falta configurar OUTLOOK_APP_PASSWORD en Vercel."
      });
    }

    const transporter = nodemailer.createTransport({
      host: "smtp.office365.com",
      port: 587,
      secure: false,
      auth: { user, pass },
      tls: { ciphers: "SSLv3" }
    });

    const files = attachments
      .filter((a) => a && a.name && a.dataUrl)
      .map((a) => {
        const base64 = String(a.dataUrl).split(",")[1] || "";
        return {
          filename: a.name,
          content: Buffer.from(base64, "base64"),
          contentType: a.type || "application/octet-stream"
        };
      });

    await transporter.sendMail({
      from: `"GPSRUTA Cobranza" <${user}>`,
      to,
      cc: user,
      subject,
      text: body,
      attachments: files
    });

    return res.status(200).json({ ok: true });
  } catch (error) {
    return res.status(500).json({
      ok: false,
      error: error?.message || "Error al enviar correo."
    });
  }
}
