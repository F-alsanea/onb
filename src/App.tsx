import React, { useState, useMemo } from 'react';
import {
  Mail,
  Code,
  Eye,
  Users,
  Download,
  CheckCircle2,
  XCircle,
  Copy,
  Smartphone,
  Monitor,
  Layout,
  Type,
  Send,
  FileText,
  Loader2,
  Lock
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

// --- Constants & Types ---

interface Recipient {
  name: string;
  email: string;
}

interface EmailConfig {
  headline: string;
  body: string;
  employeePhotoUrl: string;
  signatureName: string;
  signatureTitle: string;
  signatureTel: string;
  signatureMob: string;
  signatureEmail: string;
  signatureAddress: string;
  logoUrl: string;
}

const DEFAULT_CONFIG: EmailConfig = {
  headline: "مرحباً بك في مستقبل الأعمال الذكية",
  body: "نحن سعداء جداً بتواصلنا معك. في شركتنا، نسعى دائماً لتقديم أفضل الحلول التقنية التي تساعدك على تنمية أعمالك بكفاءة واحترافية عالية.",
  employeePhotoUrl: "https://picsum.photos/seed/employee/400/400",
  signatureName: "Faisal Alsanea",
  signatureTitle: "HR Specialist | Talent Acquisition | KAKI GROUP",
  signatureTel: "+966 (02) 6130264 – Ext 116",
  signatureMob: "+966 596995687",
  signatureEmail: "falsuni@kakigroup.co",
  signatureAddress: "Alhamraa dist., PO Box 18833, Jeddah 21425, Kingdom of Saudi Arabia",
  logoUrl: "https://kakihg.net/wp-content/uploads/2025/08/kaki_logo-footer.png"
};

const WELCOME_PRESET: Partial<EmailConfig> = {
  headline: "أهلاً بك في عائلة مجموعة كعكي",
  body: "نحن متحمسون جداً لانضمامك إلينا كعضو جديد في فريقنا. نؤمن بأن مهاراتك وخبراتك ستكون إضافة قيمة لمجموعتنا، ونتطلع للعمل معك لتحقيق نجاحات مشتركة.",
  employeePhotoUrl: "https://picsum.photos/seed/welcome/400/400",
  logoUrl: "https://kakihg.net/wp-content/uploads/2025/08/kaki_logo-footer.png"
};

const INTERNAL_PRESET: Partial<EmailConfig> = {
  headline: "تحديث داخلي لفريق العمل",
  body: "الزملاء الأعزاء، نود مشاركتكم هذا التحديث الهام بخصوص سير العمل والمستجدات الأخيرة في القسم. نشكر لكم جهودكم المستمرة.",
  employeePhotoUrl: "",
  logoUrl: ""
};

// --- Helper Components ---

const TabButton = ({ active, onClick, icon: Icon, label }: { active: boolean, onClick: () => void, icon: any, label: string }) => (
  <button
    onClick={onClick}
    className={`flex items-center gap-2 px-6 py-3 text-sm font-medium transition-all border-b-2 ${active
      ? 'border-black text-black'
      : 'border-transparent text-gray-400 hover:text-gray-600'
      }`}
  >
    <Icon size={18} />
    {label}
  </button>
);

// --- Main App ---

export default function App() {
  const [activeTab, setActiveTab] = useState<'preview' | 'editor' | 'recipients' | 'code' | 'send'>('editor');
  const [previewMode, setPreviewMode] = useState<'desktop' | 'mobile'>('desktop');
  const [config, setConfig] = useState<EmailConfig>(DEFAULT_CONFIG);
  const [recipients, setRecipients] = useState<Recipient[]>([
    { name: "هلا", email: "halbangah@kakigroup.co" },
    { name: "علي", email: "akhawaji@kakigroup.co" },
    { name: "توفيق", email: "taljuhani@kakigroup.co" },
    { name: "تركي المالكي", email: "tmalki@kakigroup.co" },
    { name: "عايض المالكي", email: "aalmalki@kakigroup.co" },
    { name: "توفيق الجهني", email: "taljuhani@kakigroup.co" },
    { name: "محمد الزهراني", email: "malzahrani@kakigroup.co" },
    { name: "دينا الغامدي", email: "dalghamdi@kakigroup.co" },
    { name: "علي خواجي", email: "akhawaji@kakigroup.co" },
    { name: "عبدالعزيز فالح", email: "abdulaziz-faleh@kakigroup.co" },
    { name: "فيصل السني", email: "falsuni@kakigroup.co" },
    { name: "روان الغامدي", email: "ralghamdi@kakigroup.co" },
    { name: "نشأت قواس", email: "ngawass@kakigroup.co" },
    { name: "رشا ساعاتي", email: "ralsaati@kakigroup.co" },
    { name: "هيفاء الجعيد", email: "haljuaid@kakigroup.co" },
    { name: "مشاري الجهني", email: "maljuhani@kakigroup.co" },
    { name: "محمد فتيني", email: "mfutaini@kakigroup.co" },
    { name: "ليبرادو", email: "buddy@kakigroup.co" },
    { name: "بندر كعكي", email: "bkaki@kakigroup.co" },
    { name: "افنان قيشاوي", email: "pr@kakigroup.co" },
    { name: "ياسين", email: "ahyasin@kakigroup.co" },
    { name: "عبدالله باقيس", email: "purchasing@kakigroup.co" },
    { name: "طه عبدالفتاح", email: "tabdufatah@kakigroup.co" },
    { name: "شاكر عاشور", email: "ashaker@kakigroup.co" },
    { name: "سيد ابراهيم", email: "sibrahim@kakigroup.co" },
    { name: "محمود اسماعيل", email: "mismael@kakigroup.co" },
    { name: "صفوت الحسيني", email: "szakouk@kakigroup.co" },
    { name: "ابرار الغامدي", email: "aghamdi@kakigroup.co" },
    { name: "فتون الردادي", email: "falradadi@kakigroup.co" },
    { name: "داره سلامه", email: "dsalama@kakigroup.co" },
    { name: "هتاف فقيه", email: "hfaqih@kakigroup.co" },
    { name: "خلود الغامدي", email: "kalghamdi@kakigroup.co" },
    { name: "ابرار الحارثي", email: "aalharthi@kakigroup.co" },
    { name: "شهد الشيخي", email: "salshaiky@kakigroup.co" },
    { name: "مروه مغربي", email: "mmagrabi@kakigroup.co" },
    { name: "وصايف ابو زاهره", email: "wabozahra@kakigroup.co" },
    { name: "هاشم المالكي", email: "halmalki@kakigroup.co" },
    { name: "احمد قاسم", email: "aqasim@kakigroup.co" },
    { name: "غسان رستم", email: "grustom@kakigroup.co" },
    { name: "احمد عبدالوهاب", email: "aabdelwahab@kakigroup.co" },
    { name: "سحر الايوبي", email: "salayoubi@kakigroup.co" },
    { name: "عبدالسلام جمعه", email: "abdulsalam@kakigroup.co" },
    { name: "عبدالعزيز صالح المالكي", email: "asaleh@kakigroup.co" },
    { name: "فادي اورفلي", email: "forfali@kakigroup.co" },
    { name: "نواف الثبيتي", email: "nalthobaity@kakigroup.co" },
    { name: "خالد صلاح", email: "ksalah@kakigroup.co" },
    { name: "رشدي جميل", email: "rjameel@kakigroup.co" },
    { name: "سامي السيد", email: "selsayd@kakigroup.co" },
    { name: "محمد ادم", email: "moadam@kakigroup.co" },
    { name: "سامح حسن", email: "shassan@kakigroup.co" }
  ]);
  const [copied, setCopied] = useState(false);
  const [previewName, setPreviewName] = useState('فيصل');

  // --- Send State ---
  const [emailPassword, setEmailPassword] = useState('');
  const [isSending, setIsSending] = useState(false);
  const [sendResults, setSendResults] = useState<Array<{ name: string; email: string; success: boolean; error?: string }>>([]);
  const [sendSummary, setSendSummary] = useState<{ successCount: number; total: number } | null>(null);
  const [sendError, setSendError] = useState('');

  // Generate the HTML Template String
  const htmlTemplate = useMemo(() => {
    return `
<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Arabic:wght@400;700&display=swap" rel="stylesheet">
    <style>
        @media only screen and (max-width: 600px) {
            .container { width: 100% !important; padding: 20px !important; }
            .headline { font-size: 28px !important; }
        }
        body, table, td, p, a, h1 { font-family: 'IBM Plex Sans Arabic', 'Tahoma', 'Arial', sans-serif !important; }
    </style>
</head>
<body style="margin: 0; padding: 0; background-color: #ffffff; font-family: 'IBM Plex Sans Arabic', 'Tahoma', 'Arial', sans-serif;">
    <center>
        <table class="container" width="600" cellpadding="0" cellspacing="0" border="0" style="width: 600px; margin: 0 auto; background-color: #ffffff;">
            <!-- Logo Section -->
            ${config.logoUrl ? `
            <tr>
                <td style="padding: 60px 0 40px 0; text-align: center;">
                    <img src="${config.logoUrl}" alt="Logo" style="max-width: 140px; height: auto; display: block; margin: 0 auto;" referrerPolicy="no-referrer">
                </td>
            </tr>
            ` : '<tr><td style="padding: 40px 0 0 0;"></td></tr>'}
            
            <!-- Headline Section -->
            <tr>
                <td style="padding: 0 40px 30px 40px; text-align: center;">
                    <h1 class="headline" style="margin: 0; color: #000000; font-size: 38px; font-weight: 700; line-height: 1.2; letter-spacing: -0.02em;">
                        ${config.headline}
                    </h1>
                </td>
            </tr>

            <!-- Greeting & Body -->
            <tr>
                <td style="padding: 20px 40px; text-align: right; color: #000000; font-size: 17px; line-height: 1.7;">
                    <p style="margin: 0 0 24px 0; font-weight: 700;">عزيزي/عزيزتي {customer_name}،</p>
                    <p style="margin: 0; font-weight: 400; opacity: 0.8;">${config.body}</p>
                </td>
            </tr>

            <!-- Employee Photo Section -->
            ${config.employeePhotoUrl ? `
            <tr>
                <td style="padding: 20px 0 30px 0; text-align: center;">
                    <img src="${config.employeePhotoUrl}" alt="Employee Photo" style="width: 100%; max-width: 600px; height: auto; display: block; margin: 0 auto;" referrerPolicy="no-referrer">
                </td>
            </tr>
            ` : ''}
            
            <!-- Signature Section -->
            <tr>
                <td style="padding: 50px 40px 30px 40px; text-align: left; border-top: 1px solid #f0f0f0; direction: ltr;">
                    <p style="margin: 0 0 15px 0; color: #000000; font-size: 16px; font-weight: 400;">Sincerely,</p>
                    <p style="margin: 0; color: #000000; font-weight: 700; font-size: 18px;">${config.signatureName}</p>
                    <p style="margin: 4px 0 12px 0; color: #000000; font-size: 14px; font-weight: 600;">${config.signatureTitle}</p>
                    
                    <div style="color: #444444; font-size: 13px; line-height: 1.6;">
                        <p style="margin: 0;">Tel. | ${config.signatureTel}</p>
                        <p style="margin: 0;">Mob: ${config.signatureMob}</p>
                        <p style="margin: 0;"><a href="#" style="color: #000000; text-decoration: underline;">Linkedin</a></p>
                        <p style="margin: 0;">E-Email: <a href="mailto:${config.signatureEmail}" style="color: #000000; text-decoration: none;">${config.signatureEmail}</a></p>
                        <p style="margin: 0;">${config.signatureAddress}</p>
                    </div>
                </td>
            </tr>
            
            <!-- Brand Logos Section -->
            <tr>
                <td style="padding: 0 40px 20px 40px; text-align: left; direction: ltr;">
                    <div style="display: block;">
                        <img src="https://kakihg.net/wp-content/uploads/2025/08/Gabbiano.png" height="35" style="height: 35px; width: auto; margin-right: 15px; margin-bottom: 10px; vertical-align: middle;" referrerPolicy="no-referrer">
                        <img src="https://kakihg.net/wp-content/uploads/2025/08/alshurafa-1.png" height="35" style="height: 35px; width: auto; margin-right: 15px; margin-bottom: 10px; vertical-align: middle;" referrerPolicy="no-referrer">
                        <img src="https://kakihg.net/wp-content/uploads/2025/08/zaikaki.png" height="35" style="height: 35px; width: auto; margin-right: 15px; margin-bottom: 10px; vertical-align: middle;" referrerPolicy="no-referrer">
                        <img src="https://kakihg.net/wp-content/uploads/2025/08/house.png" height="35" style="height: 35px; width: auto; margin-right: 15px; margin-bottom: 10px; vertical-align: middle;" referrerPolicy="no-referrer">
                        <img src="https://kakihg.net/wp-content/uploads/2025/08/gourmet-1-1.png" height="35" style="height: 35px; width: auto; margin-right: 15px; margin-bottom: 10px; vertical-align: middle;" referrerPolicy="no-referrer">
                        <img src="https://kakihg.net/wp-content/uploads/2025/08/Sezione.jpg" height="35" style="height: 35px; width: auto; margin-right: 15px; margin-bottom: 10px; vertical-align: middle;" referrerPolicy="no-referrer">
                    </div>
                </td>
            </tr>
            
            <!-- Environmental Message -->
            <tr>
                <td style="padding: 0 40px 20px 40px; text-align: left; direction: ltr;">
                    <p style="margin: 0; color: #2d8a2d; font-size: 11px; font-style: italic; font-weight: 600;">
                        🌱 Please consider the environment before printing this e-mail
                    </p>
                </td>
            </tr>
            
            <!-- Disclaimer -->
            <tr>
                <td style="padding: 0 40px 40px 40px; text-align: left; direction: ltr;">
                    <p style="margin: 0; color: #999999; font-size: 10px; line-height: 1.4; text-align: justify;">
                        <strong>Disclaimer:</strong> This message and any attachments thereto belongs to KAKI GROUP and are intended solely for the addressed recipient's and may contain confidential information. If you are not the intended recipient, please notify the sender by reply email and delete the email including any attachments thereto without producing, distributing or retaining any copies thereof. Any review, dissemination or other use of/or taking of any action in reliance upon, this information by persons or entities other than the intended recipient's is prohibited.
                    </p>
                </td>
            </tr>
            
            <!-- Footer -->
            <tr>
                <td style="padding: 20px 40px 40px 40px; text-align: center; color: #bbbbbb; font-size: 12px;">
                    <p style="margin: 0;">&copy; ${new Date().getFullYear()} جميع الحقوق محفوظة.</p>
                    <p style="margin: 8px 0 0 0;">
                        <a href="https://kakihg.net/ar/" style="color: #bbbbbb; text-decoration: underline;">kakihg.net</a>
                    </p>
                </td>
            </tr>
        </table>
    </center>
</body>
</html>
    `.trim();
  }, [config]);

  // Generate Python Code
  const pythonCode = useMemo(() => {
    return `
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time

# --- الإعدادات (Settings) ---
# بريد الشركة (KAKI GROUP)
SENDER_EMAIL = "falsuni@kakigroup.co"
# كلمة مرور التطبيق (App Password)
# إذا كنت تستخدم Microsoft 365: اذهب لـ https://mysignins.microsoft.com/security-info وأنشئ App Password
# إذا كنت تستخدم بريد عادي: استخدم كلمة مرور حسابك مباشرة
SENDER_PASSWORD = "your_app_password_here"

# --- إعدادات SMTP ---
# Microsoft 365 / Outlook
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
# إذا كان السيرفر مختلف، جرب أحد الخيارات التالية:
# SMTP_SERVER = "mail.kakigroup.co"    # سيرفر الشركة المباشر
# SMTP_SERVER = "smtp.gmail.com"       # Gmail
# SMTP_PORT = 465  # لـ SSL بدل TLS

# --- القالب (Template) ---
HTML_TEMPLATE = """
${htmlTemplate}
"""

def send_emails(excel_file_path):
    try:
        # قراءة بيانات الموظفين من ملف Excel
        # يجب أن يحتوي الملف على أعمدة باسم 'name' و 'email'
        df = pd.read_excel(excel_file_path)
        
        # الاتصال بخادم SMTP
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        
        total = len(df)
        success = 0
        
        for index, row in df.iterrows():
            customer_name = row['name']
            customer_email = row['email']
            
            # تخصيص القالب لكل موظف
            personalized_html = HTML_TEMPLATE.replace("{customer_name}", customer_name)
            
            # إنشاء الرسالة
            msg = MIMEMultipart('alternative')
            msg['From'] = f"Faisal Alsanea <{SENDER_EMAIL}>"
            msg['To'] = customer_email
            msg['Subject'] = "${config.headline}"
            
            msg.attach(MIMEText(personalized_html, 'html', 'utf-8'))
            
            # إرسال البريد
            server.send_message(msg)
            success += 1
            print(f"✅ [{success}/{total}] تم الإرسال إلى: {customer_name} ({customer_email})")
            
            # تأخير بسيط لتجنب حظر السيرفر
            time.sleep(2)
            
        server.quit()
        print(f"\\\\n✨ تمت عملية الإرسال بنجاح! ({success}/{total})")
        
    except Exception as e:
        print(f"❌ حدث خطأ: {e}")

# لتشغيل الكود:
# 1. تأكد من تثبيت المكتبات: pip install pandas openpyxl
# 2. ضع مسار ملف الإكسل هنا
# send_emails("employees.xlsx")
    `.trim();
  }, [htmlTemplate, config.headline]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const rawData = XLSX.utils.sheet_to_json(ws) as any[];

        if (rawData.length === 0) {
          alert("الملف فارغ أو لا يحتوي على بيانات صحيحة.");
          return;
        }

        const newRecipients = rawData.map(item => {
          // Try to find name in various column names
          const name = item.name || item.Name || item['الاسم'] || item['الاسم الكامل'] || item['Full Name'] || item['اسم العميل'] || 'عميل';
          // Try to find email in various column names
          const email = item.email || item.Email || item['الإيميل'] || item['البريد'] || item['البريد الإلكتروني'] || item['Email Address'] || item['عنوان البريد'];

          return { name, email };
        }).filter(r => r.email && typeof r.email === 'string' && r.email.includes('@'));

        if (newRecipients.length === 0) {
          alert("لم يتم العثور على عناوين بريد إلكتروني صحيحة. تأكد من أن أعمدة الإكسل تحتوي على عناوين مثل 'Email' أو 'الإيميل'.");
          return;
        }

        setRecipients(newRecipients);
      } catch (error) {
        console.error("Excel parsing error:", error);
        alert("حدث خطأ أثناء قراءة ملف الإكسل. تأكد من أنه ملف صحيح.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handlePhotoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // Check if it's an image
    if (!file.type.startsWith('image/')) {
      alert('يرجى اختيار ملف صورة (JPG, PNG, etc.)');
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      const base64 = evt.target?.result as string;
      setConfig({ ...config, employeePhotoUrl: base64 });
    };
    reader.readAsDataURL(file);
  };

  const copyToClipboard = () => {
    navigator.clipboard.writeText(pythonCode);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const downloadForWord = () => {
    // Word handles HTML better if we provide specific MSO headers and clean up the structure
    const header = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' 
            xmlns:w='urn:schemas-microsoft-com:office:word' 
            xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>Email Design</title>
        <!--[if gte mso 9]>
        <xml>
          <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>100</w:Zoom>
            <w:DoNotOptimizeForBrowser/>
          </w:WordDocument>
        </xml>
        <![endif]-->
        <style>
          /* Ensure Arabic fonts render correctly in Word */
          body, table, td, p, a, h1 { 
            font-family: 'Tahoma', 'Arial', sans-serif !important; 
          }
        </style>
      </head>
      <body lang=AR-SA style='tab-interval:36.0pt'>
    `;
    const footer = "</body></html>";

    // Extract only the content inside the <body> of our template to avoid nested <html> tags
    const bodyMatch = htmlTemplate.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    const bodyContent = bodyMatch ? bodyMatch[1] : htmlTemplate;

    const sourceHTML = header + bodyContent.replace("{customer_name}", previewName) + footer;

    const blob = new Blob(['\ufeff', sourceHTML], {
      type: 'application/msword'
    });

    saveAs(blob, `Email_Design_${previewName}.doc`);
  };

  const handleSendEmails = async () => {
    if (!emailPassword.trim()) {
      setSendError('يرجى إدخال كلمة مرور البريد الإلكتروني');
      return;
    }
    if (recipients.length === 0) {
      setSendError('لا يوجد مستلمون. أضف مستلمين أولاً من تاب المستلمون');
      return;
    }
    setIsSending(true);
    setSendResults([]);
    setSendSummary(null);
    setSendError('');

    try {
      const response = await fetch('/api/send-emails', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          recipients,
          htmlTemplate,
          subject: config.headline,
          password: emailPassword,
        }),
      });

      const data = await response.json();

      if (!response.ok) {
        setSendError(data.error || 'حدث خطأ غير متوقع');
        return;
      }

      setSendResults(data.results);
      setSendSummary({ successCount: data.successCount, total: data.total });
    } catch (err: any) {
      setSendError(`لم يتمكن من الاتصال بالسيرفر. تأكد من تشغيل الأمر: npm run start`);
    } finally {
      setIsSending(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#F9F9F9] text-[#1A1A1A] font-sans selection:bg-black selection:text-white">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-black rounded-lg flex items-center justify-center">
              <Mail className="text-white" size={18} />
            </div>
            <h1 className="font-semibold tracking-tight text-lg">ProMail Designer</h1>
          </div>

          <div className="flex gap-2">
            <button
              onClick={downloadForWord}
              className="flex items-center gap-2 px-4 py-2 bg-white border border-gray-200 text-black rounded-full text-sm font-medium hover:bg-gray-50 transition-colors"
            >
              <FileText size={16} />
              تحميل لـ Word
            </button>
            <button
              onClick={copyToClipboard}
              className="flex items-center gap-2 px-4 py-2 bg-black text-white rounded-full text-sm font-medium hover:bg-gray-800 transition-colors"
            >
              {copied ? <CheckCircle2 size={16} /> : <Copy size={16} />}
              {copied ? 'تم النسخ' : 'نسخ كود Python'}
            </button>
          </div>
        </div>

        <div className="max-w-7xl mx-auto px-6 flex overflow-x-auto no-scrollbar">
          <TabButton active={activeTab === 'editor'} onClick={() => setActiveTab('editor')} icon={Layout} label="المحرر" />
          <TabButton active={activeTab === 'preview'} onClick={() => setActiveTab('preview')} icon={Eye} label="المعاينة" />
          <TabButton active={activeTab === 'recipients'} onClick={() => setActiveTab('recipients')} icon={Users} label="المستلمون" />
          <TabButton active={activeTab === 'code'} onClick={() => setActiveTab('code')} icon={Code} label="كود Python" />
          <TabButton active={activeTab === 'send'} onClick={() => setActiveTab('send')} icon={Send} label="إرسال" />
        </div>
      </header>

      <main className="max-w-7xl mx-auto p-6 md:p-10">
        <AnimatePresence mode="wait">
          {activeTab === 'editor' && (
            <motion.div
              key="editor"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="grid grid-cols-1 lg:grid-cols-2 gap-10"
            >
              <div className="space-y-8">
                <section>
                  <h2 className="text-xs font-bold uppercase tracking-widest text-gray-400 mb-4 flex items-center gap-2">
                    <Layout size={14} /> اختيار القالب
                  </h2>
                  <div className="flex gap-3">
                    <button
                      onClick={() => setConfig({ ...config, ...DEFAULT_CONFIG })}
                      className="flex-1 px-4 py-3 rounded-xl border border-gray-200 text-sm font-medium hover:border-black transition-all bg-white"
                    >
                      حملة تسويقية
                    </button>
                    <button
                      onClick={() => setConfig({ ...config, ...WELCOME_PRESET })}
                      className="flex-1 px-4 py-3 rounded-xl border border-gray-200 text-sm font-medium hover:border-black transition-all bg-white"
                    >
                      ترحيب بموظف جديد
                    </button>
                    <button
                      onClick={() => setConfig({ ...config, ...INTERNAL_PRESET })}
                      className="flex-1 px-4 py-3 rounded-xl border border-gray-200 text-sm font-medium hover:border-black transition-all bg-white"
                    >
                      مراسلات داخلية
                    </button>
                  </div>
                </section>

                <section>
                  <h2 className="text-xs font-bold uppercase tracking-widest text-gray-400 mb-4 flex items-center gap-2">
                    <Type size={14} /> محتوى الرسالة
                  </h2>
                  <div className="space-y-4">
                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-1">اسم الشخص (للمعاينة)</label>
                      <input
                        type="text"
                        value={previewName}
                        onChange={(e) => setPreviewName(e.target.value)}
                        placeholder="مثال: أحمد"
                        className="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-black focus:ring-0 transition-all outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-1">العنوان الرئيسي</label>
                      <input
                        type="text"
                        value={config.headline}
                        onChange={(e) => setConfig({ ...config, headline: e.target.value })}
                        className="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-black focus:ring-0 transition-all outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-gray-700 mb-1">نص الرسالة</label>
                      <textarea
                        rows={5}
                        value={config.body}
                        onChange={(e) => setConfig({ ...config, body: e.target.value })}
                        className="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-black focus:ring-0 transition-all outline-none resize-none"
                      />
                    </div>
                  </div>
                </section>

                <section>
                  <h2 className="text-xs font-bold uppercase tracking-widest text-gray-400 mb-4">صورة الموظف</h2>
                  <div className="space-y-4">
                    <div className="flex gap-2">
                      <input
                        type="text"
                        placeholder="رابط صورة الموظف (URL)"
                        value={config.employeePhotoUrl.startsWith('data:') ? 'صورة مرفوعة محلياً' : config.employeePhotoUrl}
                        onChange={(e) => setConfig({ ...config, employeePhotoUrl: e.target.value })}
                        disabled={config.employeePhotoUrl.startsWith('data:')}
                        className="flex-1 px-4 py-2 rounded-lg border border-gray-200 focus:border-black outline-none text-sm disabled:bg-gray-50 disabled:text-gray-400"
                      />
                      {config.employeePhotoUrl.startsWith('data:') && (
                        <button
                          onClick={() => setConfig({ ...config, employeePhotoUrl: '' })}
                          className="px-3 py-2 bg-red-50 text-red-600 rounded-lg text-xs font-medium hover:bg-red-100 transition-colors"
                        >
                          حذف
                        </button>
                      )}
                    </div>

                    <div className="relative">
                      <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed border-gray-200 rounded-2xl cursor-pointer hover:bg-gray-50 hover:border-black transition-all group">
                        <div className="flex flex-col items-center justify-center pt-5 pb-6">
                          <Download className="w-8 h-8 text-gray-300 group-hover:text-black mb-2 transition-colors" />
                          <p className="text-sm text-gray-500 font-medium">رفع صورة من الجهاز</p>
                          <p className="text-[10px] text-gray-400 mt-1">JPG, PNG (سيتم تحويلها لـ Base64)</p>
                        </div>
                        <input type="file" accept="image/*" className="hidden" onChange={handlePhotoUpload} />
                      </label>
                    </div>

                    <p className="text-[10px] text-gray-400 leading-relaxed">
                      * ملاحظة: الصور المرفوعة محلياً يتم تضمينها كـ Base64. في رسائل البريد الحقيقية، يفضل استخدام روابط (URL) لتقليل حجم الرسالة وتجنب الحظر.
                    </p>
                  </div>
                </section>

                <section className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">الاسم الكامل</label>
                    <input
                      type="text"
                      value={config.signatureName}
                      onChange={(e) => setConfig({ ...config, signatureName: e.target.value })}
                      className="w-full px-4 py-2 rounded-lg border border-gray-200 focus:border-black outline-none text-sm"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">المسمى الوظيفي والشركة</label>
                    <input
                      type="text"
                      value={config.signatureTitle}
                      onChange={(e) => setConfig({ ...config, signatureTitle: e.target.value })}
                      className="w-full px-4 py-2 rounded-lg border border-gray-200 focus:border-black outline-none text-sm"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">رقم الهاتف (Tel)</label>
                    <input
                      type="text"
                      value={config.signatureTel}
                      onChange={(e) => setConfig({ ...config, signatureTel: e.target.value })}
                      className="w-full px-4 py-2 rounded-lg border border-gray-200 focus:border-black outline-none text-sm"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">رقم الجوال (Mob)</label>
                    <input
                      type="text"
                      value={config.signatureMob}
                      onChange={(e) => setConfig({ ...config, signatureMob: e.target.value })}
                      className="w-full px-4 py-2 rounded-lg border border-gray-200 focus:border-black outline-none text-sm"
                    />
                  </div>
                  <div className="md:col-span-2">
                    <label className="block text-sm font-medium text-gray-700 mb-1">البريد الإلكتروني</label>
                    <input
                      type="text"
                      value={config.signatureEmail}
                      onChange={(e) => setConfig({ ...config, signatureEmail: e.target.value })}
                      className="w-full px-4 py-2 rounded-lg border border-gray-200 focus:border-black outline-none text-sm"
                    />
                  </div>
                  <div className="md:col-span-2">
                    <label className="block text-sm font-medium text-gray-700 mb-1">العنوان</label>
                    <input
                      type="text"
                      value={config.signatureAddress}
                      onChange={(e) => setConfig({ ...config, signatureAddress: e.target.value })}
                      className="w-full px-4 py-2 rounded-lg border border-gray-200 focus:border-black outline-none text-sm"
                    />
                  </div>
                </section>
              </div>

              <div className="hidden lg:block bg-white rounded-3xl border border-gray-100 shadow-sm overflow-hidden sticky top-32 h-fit">
                <div className="p-4 border-b border-gray-50 flex justify-between items-center">
                  <span className="text-xs font-semibold text-gray-400">معاينة سريعة</span>
                  <div className="flex gap-2">
                    <button onClick={() => setPreviewMode('desktop')} className={`p-1.5 rounded ${previewMode === 'desktop' ? 'bg-gray-100' : ''}`}><Monitor size={14} /></button>
                    <button onClick={() => setPreviewMode('mobile')} className={`p-1.5 rounded ${previewMode === 'mobile' ? 'bg-gray-100' : ''}`}><Smartphone size={14} /></button>
                  </div>
                </div>
                <div className="p-4 bg-gray-50 flex justify-center">
                  <div className={`bg-white shadow-2xl transition-all duration-500 overflow-auto no-scrollbar ${previewMode === 'mobile' ? 'w-[320px] h-[500px]' : 'w-full h-[500px]'}`}>
                    <iframe
                      title="Email Preview"
                      srcDoc={htmlTemplate.replace("{customer_name}", previewName)}
                      className="w-full h-full border-none"
                    />
                  </div>
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'preview' && (
            <motion.div
              key="preview"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="flex flex-col items-center"
            >
              <div className="mb-8 flex gap-4 bg-white p-2 rounded-full border border-gray-200 shadow-sm">
                <button
                  onClick={() => setPreviewMode('desktop')}
                  className={`flex items-center gap-2 px-6 py-2 rounded-full text-sm font-medium transition-all ${previewMode === 'desktop' ? 'bg-black text-white' : 'text-gray-500 hover:bg-gray-50'}`}
                >
                  <Monitor size={16} /> سطح المكتب
                </button>
                <button
                  onClick={() => setPreviewMode('mobile')}
                  className={`flex items-center gap-2 px-6 py-2 rounded-full text-sm font-medium transition-all ${previewMode === 'mobile' ? 'bg-black text-white' : 'text-gray-500 hover:bg-gray-50'}`}
                >
                  <Smartphone size={16} /> الجوال
                </button>
              </div>

              <div className={`bg-white shadow-2xl transition-all duration-500 rounded-2xl overflow-hidden ${previewMode === 'mobile' ? 'w-[375px]' : 'w-full max-w-[800px]'}`}>
                <iframe
                  title="Full Preview"
                  srcDoc={htmlTemplate.replace("{customer_name}", previewName)}
                  className="w-full h-[800px] border-none"
                />
              </div>
            </motion.div>
          )}

          {activeTab === 'recipients' && (
            <motion.div
              key="recipients"
              initial={{ opacity: 0, scale: 0.98 }}
              animate={{ opacity: 1, scale: 1 }}
              className="max-w-3xl mx-auto"
            >
              <div className="bg-white rounded-3xl border border-gray-200 overflow-hidden shadow-sm">
                <div className="p-8 border-b border-gray-100 flex justify-between items-center">
                  <div>
                    <h2 className="text-xl font-bold">قائمة المستلمين</h2>
                    <p className="text-sm text-gray-500 mt-1">قم برفع ملف Excel أو إضافة الأسماء يدوياً</p>
                  </div>
                  <label className="cursor-pointer bg-gray-50 hover:bg-gray-100 text-black px-6 py-3 rounded-2xl text-sm font-semibold flex items-center gap-2 transition-all border border-gray-200">
                    <Download size={18} />
                    رفع Excel
                    <input type="file" accept=".xlsx, .xls" className="hidden" onChange={handleFileUpload} />
                  </label>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-right">
                    <thead className="bg-gray-50 text-xs font-bold text-gray-400 uppercase tracking-widest">
                      <tr>
                        <th className="px-8 py-4">الاسم</th>
                        <th className="px-8 py-4">البريد الإلكتروني</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-50">
                      {recipients.map((r, i) => (
                        <tr key={i} className="hover:bg-gray-50/50 transition-colors">
                          <td className="px-8 py-4 font-medium">{r.name}</td>
                          <td className="px-8 py-4 text-gray-500 font-mono text-sm">{r.email}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {recipients.length === 0 && (
                  <div className="p-20 text-center text-gray-400">
                    <Users size={48} className="mx-auto mb-4 opacity-20" />
                    <p>لا يوجد مستلمون حالياً</p>
                  </div>
                )}
              </div>
            </motion.div>
          )}

          {activeTab === 'code' && (
            <motion.div
              key="code"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="max-w-4xl mx-auto"
            >
              <div className="bg-[#1E1E1E] rounded-3xl overflow-hidden shadow-2xl border border-white/5">
                <div className="px-6 py-4 bg-white/5 border-b border-white/5 flex justify-between items-center">
                  <div className="flex gap-1.5">
                    <div className="w-3 h-3 rounded-full bg-[#FF5F56]"></div>
                    <div className="w-3 h-3 rounded-full bg-[#FFBD2E]"></div>
                    <div className="w-3 h-3 rounded-full bg-[#27C93F]"></div>
                  </div>
                  <span className="text-xs font-mono text-gray-500">email_sender.py</span>
                  <button
                    onClick={copyToClipboard}
                    className="text-xs text-gray-400 hover:text-white flex items-center gap-1.5 transition-colors"
                  >
                    {copied ? <CheckCircle2 size={14} className="text-green-400" /> : <Copy size={14} />}
                    {copied ? 'تم النسخ' : 'نسخ الكود'}
                  </button>
                </div>
                <div className="p-8 overflow-x-auto">
                  <pre className="text-sm font-mono text-gray-300 leading-relaxed">
                    <code>{pythonCode}</code>
                  </pre>
                </div>
              </div>

              <div className="mt-8 grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="bg-white p-6 rounded-2xl border border-gray-200">
                  <h3 className="font-bold mb-2 flex items-center gap-2">
                    <CheckCircle2 size={18} className="text-green-500" /> متطلبات التشغيل
                  </h3>
                  <ul className="text-sm text-gray-600 space-y-2 list-disc list-inside">
                    <li>تثبيت بايثون على جهازك</li>
                    <li>تثبيت المكتبات: <code className="bg-gray-100 px-1 rounded">pip install pandas openpyxl</code></li>
                    <li>معرفة سيرفر SMTP لبريد الشركة</li>
                    <li>إنشاء <span className="font-semibold">App Password</span> من إعدادات Microsoft 365</li>
                  </ul>
                </div>
                <div className="bg-black text-white p-6 rounded-2xl">
                  <h3 className="font-bold mb-2 flex items-center gap-2">
                    <Send size={18} className="text-blue-400" /> كيف يعمل؟
                  </h3>
                  <p className="text-sm text-gray-400 leading-relaxed">
                    السكربت يقوم بقراءة ملف Excel، ويستبدل المتغير <code className="text-white">{`{customer_name}`}</code> داخل القالب باسم كل عميل، ثم يرسل الإيميل بشكل آلي واحترافي.
                  </p>
                </div>
              </div>
            </motion.div>
          )}

          {/* ===== SEND TAB ===== */}
          {activeTab === 'send' && (
            <motion.div
              key="send"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="max-w-2xl mx-auto"
            >
              {/* Info Card */}
              <div className="bg-white rounded-3xl border border-gray-200 p-8 shadow-sm mb-6">
                <div className="flex items-center gap-3 mb-6">
                  <div className="w-10 h-10 bg-black rounded-xl flex items-center justify-center">
                    <Send size={18} className="text-white" />
                  </div>
                  <div>
                    <h2 className="font-bold text-lg">إرسال الرسائل</h2>
                    <p className="text-sm text-gray-500">سيتم إرسال القالب كاملاً لـ {recipients.length} مستلم</p>
                  </div>
                </div>

                {/* Sender Info */}
                <div className="bg-gray-50 rounded-2xl p-4 mb-6 flex items-center gap-3">
                  <Mail size={16} className="text-gray-400" />
                  <div>
                    <p className="text-xs text-gray-400">المُرسِل</p>
                    <p className="text-sm font-semibold">Faisal Alsanea &lt;falsuni@kakigroup.co&gt;</p>
                  </div>
                </div>

                {/* Password Field */}
                <div className="mb-6">
                  <label className="block text-sm font-medium text-gray-700 mb-2 flex items-center gap-2">
                    <Lock size={14} /> كلمة مرور بريد الشركة (App Password)
                  </label>
                  <input
                    type="password"
                    value={emailPassword}
                    onChange={(e) => setEmailPassword(e.target.value)}
                    placeholder="أدخل كلمة مرور بريدك الإلكتروني..."
                    className="w-full px-4 py-3 rounded-xl border border-gray-200 focus:border-black focus:ring-0 transition-all outline-none font-mono"
                    disabled={isSending}
                  />
                  <p className="text-xs text-gray-400 mt-2">كلمة المرور لا تُحفظ ولا تُرسل إلا للسيرفر المحلي</p>
                </div>

                {/* Error Message */}
                {sendError && (
                  <div className="bg-red-50 border border-red-100 text-red-700 rounded-xl p-4 mb-4 text-sm">
                    ❌ {sendError}
                  </div>
                )}

                {/* Send Button */}
                <button
                  onClick={handleSendEmails}
                  disabled={isSending}
                  className="w-full flex items-center justify-center gap-3 px-6 py-4 bg-black text-white rounded-2xl font-bold text-base hover:bg-gray-800 transition-colors disabled:opacity-60 disabled:cursor-not-allowed"
                >
                  {isSending ? (
                    <><Loader2 size={20} className="animate-spin" /> جاري الإرسال...</>
                  ) : (
                    <><Send size={20} /> إرسال لجميع الموظفين ({recipients.length})</>
                  )}
                </button>
              </div>

              {/* Summary */}
              {sendSummary && (
                <div className={`rounded-2xl p-5 mb-6 flex items-center gap-3 ${sendSummary.successCount === sendSummary.total
                    ? 'bg-green-50 border border-green-100'
                    : 'bg-yellow-50 border border-yellow-100'
                  }`}>
                  <CheckCircle2 size={22} className="text-green-600 shrink-0" />
                  <p className="font-bold text-green-800">
                    ✨ تمت العملية: {sendSummary.successCount} / {sendSummary.total} تم الإرسال بنجاح
                  </p>
                </div>
              )}

              {/* Results List */}
              {sendResults.length > 0 && (
                <div className="bg-white rounded-3xl border border-gray-200 overflow-hidden shadow-sm">
                  <div className="px-6 py-4 border-b border-gray-100">
                    <h3 className="font-bold text-sm">تفاصيل الإرسال</h3>
                  </div>
                  <div className="divide-y divide-gray-50 max-h-[400px] overflow-y-auto">
                    {sendResults.map((r, i) => (
                      <div key={i} className="px-6 py-4 flex items-center justify-between">
                        <div>
                          <p className="font-medium text-sm">{r.name}</p>
                          <p className="text-xs text-gray-400 font-mono">{r.email}</p>
                          {r.error && <p className="text-xs text-red-500 mt-1">{r.error}</p>}
                        </div>
                        {r.success
                          ? <CheckCircle2 size={20} className="text-green-500 shrink-0" />
                          : <XCircle size={20} className="text-red-400 shrink-0" />
                        }
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      <footer className="max-w-7xl mx-auto px-6 py-10 border-t border-gray-100 text-center">
        <p className="text-xs text-gray-400 font-medium tracking-widest uppercase">
          صمم بكل حب لخدمة أعمالكم الاحترافية
        </p>
      </footer>
    </div>
  );
}
