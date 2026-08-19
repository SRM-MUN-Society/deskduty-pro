# DeskDuty Pro

DeskDuty Pro is a smart, client-side desk duty roster generator built for the **SRM MUN Society (Internal Affairs)**. It takes availability responses from Google Forms (via Excel/CSV) and automatically schedules both Heads and Members across hourly time slots.

## Features

- **Client-Side Only**: Runs entirely in the browser with no backend required. All data processing and generation happen locally.
- **Fairness-Based Scheduling**: Intelligently tracks duty allocations to ensure duties are distributed as fairly and evenly as possible.
- **Duty Caps**: Automatically caps the maximum number of duties based on an individual's reported free hours.
- **Persistent Data**: Saves uploaded Excel data to `localStorage` so it persists across page reloads.
- **Clean UI**: A modern, responsive design built with Next.js and Tailwind CSS featuring dynamic output generation.

## Getting Started

1. Clone the repository and install dependencies:
   ```bash
   npm install
   ```

2. Create a `.env` file in the root directory and set the admin access password:
   ```env
   NEXT_PUBLIC_ADMIN_PASSWORD=your_password_here
   ```

3. Run the development server:
   ```bash
   npm run dev
   ```

4. Open [http://localhost:3000](http://localhost:3000) in your browser. Enter your access key to log in and start generating rosters!

## Usage

1. Log in using the configured access key.
2. Upload the **Heads** free slots Excel file.
3. Upload the **Members** free slots Excel file.
4. Select the **Day Order** from the dropdown menu.
5. Click **Generate Roster** to view the optimally scheduled desk duties for both Heads and Members.
6. Use the copy button to easily extract the results.

## Built With
- [Next.js](https://nextjs.org)
- [React](https://react.dev/)
- [Tailwind CSS](https://tailwindcss.com/)
- [XLSX](https://www.npmjs.com/package/xlsx) for Excel parsing
