import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: '评鉴 · 本地在线评价工作台',
  description: '自定义评价表、分配评价任务并汇总评分结果。',
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="zh-CN"><body>{children}</body></html>;
}
