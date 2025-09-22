const nodemailer = require('nodemailer');

function createTransporter(smtpOptions = {}) {
  const { host, port, secure, auth, ...rest } = smtpOptions || {};

  if (!host || typeof host !== 'string') {
    throw new Error("Le serveur SMTP (host) est requis pour créer un transporteur.");
  }

  const transporter = nodemailer.createTransport({
    host: host.trim(),
    port: port !== undefined ? Number(port) : 587,
    secure: Boolean(secure),
    auth:
      auth && typeof auth === 'object' && auth.user
        ? { user: auth.user, pass: auth.pass || '' }
        : undefined,
    ...rest,
  });

  return transporter;
}

function formatAddress(entry) {
  if (typeof entry === 'string') {
    const email = entry.trim();
    return email ? email : null;
  }

  if (!entry || typeof entry !== 'object') {
    return null;
  }

  const email = typeof entry.email === 'string' ? entry.email.trim() : '';
  if (!email) {
    return null;
  }

  const name = typeof entry.name === 'string' ? entry.name.trim() : '';
  return name ? { name, address: email } : email;
}

async function sendCampaignEmails(options = {}) {
  const {
    transporter,
    smtp,
    sender,
    recipients,
    subject,
    html,
    text,
    attachments,
    headers,
  } = options;

  let activeTransporter = transporter;
  if (!activeTransporter) {
    if (!smtp) {
      throw new Error(
        'Un transporteur Nodemailer ou des paramètres SMTP doivent être fournis pour envoyer des mails.',
      );
    }
    activeTransporter = createTransporter(smtp);
  }

  const formattedSender = formatAddress(sender);
  if (!formattedSender) {
    throw new Error("L'expéditeur doit comporter une adresse mail valide.");
  }

  const formattedRecipients = [];
  const seenAddresses = new Set();

  if (Array.isArray(recipients)) {
    recipients.forEach((recipient) => {
      const formatted = formatAddress(recipient);
      if (!formatted) {
        return;
      }

      const email = typeof formatted === 'string' ? formatted : formatted.address;
      const key = email.toLowerCase();
      if (seenAddresses.has(key)) {
        return;
      }
      seenAddresses.add(key);
      formattedRecipients.push(formatted);
    });
  }

  if (formattedRecipients.length === 0) {
    throw new Error('Aucun destinataire valide fourni.');
  }

  const message = {
    from: formattedSender,
    to: formattedRecipients,
    subject: subject && subject.toString().trim() ? subject.toString().trim() : 'Campagne UManager',
  };

  if (typeof text === 'string' && text.trim()) {
    message.text = text;
  }

  if (typeof html === 'string' && html.trim()) {
    message.html = html;
  }

  if (Array.isArray(attachments) && attachments.length > 0) {
    message.attachments = attachments
      .filter((item) => item && typeof item === 'object')
      .map((item) => ({ ...item }));
  }

  if (headers && typeof headers === 'object') {
    message.headers = { ...headers };
  }

  return activeTransporter.sendMail(message);
}

module.exports = {
  createTransporter,
  sendCampaignEmails,
  formatAddress,
};
