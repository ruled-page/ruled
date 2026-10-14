import '@ruled/sheet/style.css';

export const metadata = { title: 'Ruled in Next.js' };

export default function Layout({ children }) {
  return (
    <html lang="en">
      <body style={{ margin: 0 }}>{children}</body>
    </html>
  );
}
