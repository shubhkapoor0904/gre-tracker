# TargetMS — GRE / MS University Application Deadline Tracker

TargetMS is a complete, polished, and **100% FREE** personal application-management dashboard for tracking graduate school (MS/PhD) applications, GRE requirements, document checklists, deadlines, calendar reminders, and direct email notifications.

---

## 🌟 Key Features

1. **User Profile & Free Tier Email Dashboard**:
   - Header features a live user profile pill (`👤 Sarah Jenkins (sarah.gre2026@gmail.com)`).
   - Integrated **EmailJS SDK** for sending up to **200 completely FREE emails per month** with zero server hosting costs and 0 paid subscriptions required!
   - Features **1-Click "📧 Email" Alert Buttons** on every university card and detail drawer to trigger formatted reminder emails instantly.
   - Built-in **"🚀 Send Test Email Now"** feature in settings for immediate verification.

2. **Intelligent Excel Import (`Colleges.xlsx`)**:
   - Automatically attempts to load and parse `Colleges.xlsx` on first load using SheetJS.
   - Includes a drag-and-drop file uploader to import any updated `.xlsx`, `.xls`, or `.csv` file.
   - Dynamic header detection automatically maps university, program, degree, country, deadline, status, fee, and notes columns.
   - Preserves multiple programs for the same university as separate, independent application records.

3. **Executive SaaS Dashboard & Live Countdown**:
   - **Metrics Bar**: Instant stats on Total Applications, Submitted, In Progress, Unstarted, Critical/Overdue (<14 days), and High Priority.
   - **Timeline Carousel**: Chronologically sorted upcoming deadlines with real-time countdown pills (`82d SAFE`, `12d URGENT`, `3d CRITICAL`, `DUE TODAY`, `OVERDUE`).
   - **Modern SaaS Aesthetics**: Deep dark glassmorphism, Inter/Outfit typography, clean visual hierarchy, and gold priority accents.

4. **Dynamic Requirements Checklist & Smart Math**:
   - Tracks 11 core application requirements: Resume/CV, SOP, LOR 1-3, Transcripts, GRE, TOEFL/IELTS, App Form, App Fee, Final Submission.
   - **Smart Math Logic**: If a program waives or does not require GRE, the GRE requirement is automatically excluded from the completion % denominator so it doesn't penalize your progress.
   - Custom document items can be added per program (e.g. Portfolio, Video Essay, Diversity Statement).

5. **Smart Warnings Engine**:
   - Automated alerts banner for:
     - Deadline in $\le 14$ days with incomplete SOP or LORs.
     - Deadline in $\le 14$ days with status "Not Started".
     - Application marked "Ready to Apply" with $< 80\%$ document completion.
     - Passed / overdue deadlines.

6. **100% Free Calendar Sync & Reminders**:
   - **1-Click Google Calendar Event Links**: Instantly create deadline events with title, details, and portal URLs.
   - **Downloadable `.ics` iCalendar Files**: Single application or bulk all-deadlines `.ics` export with built-in alarm triggers (30d, 14d, 7d, 3d, 1d before).
   - **Browser Desktop Notifications**: Native Web Notification API alerts for pending deadlines.

---

## ⚠️ Free Tier Quotas, Cautions & Recommended Usage Strategy

### 📊 Free Tier Service Limits

| Service / Method | Free Tier Quota | Cost | Recommended Monthly Use |
| :--- | :--- | :--- | :--- |
| **EmailJS Free API** | **200 emails / month** | **$0 / forever** | **Milestone Alerts (30d, 14d, 7d, 3d, 1d)** |
| **Resend Free API** | **3,000 emails / month** (100/day) | **$0 / forever** | **Daily / Weekly Automated Reports** |
| **Desktop Push Alerts** | **Unlimited** | **$0 / forever** | **Every time app is opened** |
| **Google Calendar Sync** | **Unlimited** | **$0 / forever** | **All Application Deadlines** |
| **`.ics` Calendar File** | **Unlimited** | **$0 / forever** | **All Application Deadlines** |

---

### 💡 How Should the User Pace Emails Throughout the Month?

#### ❌ **Avoid: Sending Raw Daily Emails for All Universities**
- If you have **12 target universities** and send daily automated emails for all of them:
  $$12 \text{ emails/day} \times 30 \text{ days} = 360 \text{ emails/month}$$
  *This would exceed the 200 emails/month cap of EmailJS's free tier.*

---

#### ✅ **Option 1: Recommended Milestone-Based Emailing (Best Practice)**
Send email alerts **ONLY when an application hits critical countdown milestones**:
- **30 days before deadline** (Initial wake-up)
- **14 days before deadline** (Urgent document check)
- **7 days before deadline** (SOP & LOR finalization)
- **3 days before deadline** (Final portal check)
- **1 day before deadline** (Submission day)

**The Math**:
For 10 universities $\times 5$ milestone emails per university = **50 total emails over the entire 4-month application season** (less than **15 emails/month**). This uses less than **8%** of your free tier limit!

---

#### ✅ **Option 2: Weekly Sunday Digest Email**
Send **1 consolidated weekly digest email** every Sunday morning summarizing upcoming deadlines for the week:
- **4 digest emails per month** total.
- Leaves 196 free emails unused every month!

---

#### ✅ **Option 3: Combine with Unlimited Free Local Alerts**
- Use **Browser Desktop Push Notifications** and **Google Calendar Sync / `.ics` alarms** for daily reminders (100% free, zero limits).
- Use **Email Alerts** for high-priority milestone warnings.

---

## 🚀 How to Run Locally

1. Simply open [`index.html`](file:///d:/gre%20tracker/index.html) in any modern web browser (Google Chrome, Microsoft Edge, Firefox, Safari).
2. Or serve locally with `npm run dev` or `npx serve .`
3. Click **"📧 Email Alerts & Profile"** or the user pill in the header to set recipient email address (`sarah.gre2026@gmail.com`).
4. Use the **"📧 Email"** buttons on any card to send deadline alerts directly to your inbox!
5. Click **"Import Excel / CSV"** at the top right of the application header to upload `Colleges.xlsx` at any time!

---

## ⚡ Deploying to Vercel (1-Click Ready)

TargetMS is 100% configured for Vercel static deployment out of the box!

### Option A: Via Vercel Dashboard (Recommended)
1. Push this repository to GitHub, GitLab, or Bitbucket.
2. Go to [Vercel Dashboard](https://vercel.com/new) and click **"Add New..." → "Project"**.
3. Import your target repository.
4. Vercel automatically detects static configuration (`vercel.json` and `package.json`). Click **Deploy**.
5. Your TargetMS Tracker is live with a global HTTPS URL!

### Option B: Via Vercel CLI
```bash
# 1. Install Vercel CLI globally
npm i -g vercel

# 2. Deploy directly from repository root
vercel

# 3. Deploy to production domain
vercel --prod
```

