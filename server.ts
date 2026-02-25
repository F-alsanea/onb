import express from 'express';
import nodemailer from 'nodemailer';
import { Request, Response } from 'express';

const app = express();
app.use(express.json({ limit: '10mb' }));

// CORS for Vite dev server
app.use((req, res, next) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') return res.sendStatus(200);
    next();
});

interface SendRequest {
    recipients: Array<{ name: string; email: string }>;
    htmlTemplate: string;
    subject: string;
    password: string;
}

interface SendResult {
    name: string;
    email: string;
    success: boolean;
    error?: string;
}

app.post('/api/send-emails', async (req: Request, res: Response) => {
    const { recipients, htmlTemplate, subject, password, senderEmail }: any = req.body;

    if (!password) {
        return res.status(400).json({ error: 'كلمة المرور مطلوبة' });
    }
    if (!senderEmail) {
        return res.status(400).json({ error: 'إيميل المُرسل مطلوب' });
    }
    if (!recipients || recipients.length === 0) {
        return res.status(400).json({ error: 'لا يوجد مستلمون' });
    }

    // Create transporter — tries Microsoft 365 first
    const transporter = nodemailer.createTransport({
        host: 'smtp.office365.com',
        port: 587,
        secure: false,
        auth: {
            user: senderEmail,
            pass: password,
        },
        tls: {
            ciphers: 'SSLv3',
            rejectUnauthorized: false,
        },
    });

    // Verify connection before sending
    try {
        await transporter.verify();
    } catch (err: any) {
        return res.status(401).json({
            error: `فشل الاتصال: ${err.message}`,
        });
    }

    const results: SendResult[] = [];

    for (const recipient of recipients) {
        // Replace placeholder with actual name
        const personalizedHtml = htmlTemplate.replace(/\{customer_name\}/g, recipient.name);

        try {
            await transporter.sendMail({
                from: `"Faisal Alsanea | KAKI GROUP" <${senderEmail}>`,
                to: recipient.email,
                subject: subject,
                html: personalizedHtml,
                // Plain text fallback
                text: `عزيزي/عزيزتي ${recipient.name}،\n\nيرجى عرض هذا البريد في بريد يدعم HTML لعرض المحتوى الكامل.`,
            });

            results.push({ name: recipient.name, email: recipient.email, success: true });
            console.log(`✅ Sent to ${recipient.name} <${recipient.email}>`);

            // Small delay to avoid rate limiting
            await new Promise((r) => setTimeout(r, 1500));
        } catch (err: any) {
            results.push({
                name: recipient.name,
                email: recipient.email,
                success: false,
                error: err.message,
            });
            console.error(`❌ Failed: ${recipient.email} — ${err.message}`);
        }
    }

    const successCount = results.filter((r) => r.success).length;
    console.log(`\n✨ Done: ${successCount}/${recipients.length} sent`);

    res.json({ results, successCount, total: recipients.length });
});

const PORT = 3002;
app.listen(PORT, () => {
    console.log(`📧 Email server running on http://localhost:${PORT}`);
});
