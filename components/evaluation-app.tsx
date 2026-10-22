'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart3,
  Check,
  ChevronRight,
  ClipboardCheck,
  Copy,
  Download,
  Eye,
  FileSpreadsheet,
  KeyRound,
  LayoutDashboard,
  LogIn,
  LogOut,
  Menu,
  PencilLine,
  Plus,
  Settings,
  Sparkles,
  Trash2,
  Upload,
  UserPlus,
  Users,
  X,
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Checkbox } from '@/components/ui/checkbox';
import { Input } from '@/components/ui/input';
import { Textarea } from '@/components/ui/textarea';
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog';

type View =
  | 'dashboard'
  | 'templates'
  | 'reviewers'
  | 'tasks'
  | 'results'
  | 'reviewer';
type Criterion = {
  id: string;
  name: string;
  dimension: string;
  weight: number;
  max: number;
  description: string;
};
type SheetCellRole =
  | 'score'
  | 'input'
  | 'comment'
  | 'remarks'
  | 'total'
  | 'subject'
  | 'reviewer'
  | 'stage'
  | 'signature';
type SheetCell = {
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
  text: string;
  role?: SheetCellRole;
  criterionId?: string;
  fieldId?: string;
  style?: {
    bold?: boolean;
    fontSize?: number;
    align?: 'left' | 'center' | 'right';
    vertical?: 'top' | 'middle' | 'bottom';
    background?: string;
    color?: string;
    wrap?: boolean;
  };
};
type SheetLayout = {
  sheetName: string;
  rowCount: number;
  colCount: number;
  columnWidths: number[];
  rowHeights: number[];
  cells: SheetCell[];
};
type Template = {
  id: string;
  name: string;
  description: string;
  criteria: Criterion[];
  updatedAt: string;
  sheetLayout?: SheetLayout;
  version?: number;
  familyId?: string;
};
type Reviewer = {
  id: string;
  name: string;
  department: string;
  code: string;
  active: boolean;
};
type Assignment = {
  id: string;
  subject: string;
  department: string;
  reviewer: string;
  reviewerId?: string;
  status: 'pending' | 'submitted';
  scores: Record<string, number>;
  comment: string;
  remarks?: string;
  fields?: Record<string, string>;
  submittedAt?: string;
};
type Task = {
  id: string;
  name: string;
  templateId: string;
  deadline: string;
  stageStartDate?: string;
  stageEndDate?: string;
  stageDate?: string;
  createdAt: string;
  assignments: Assignment[];
};
type AppData = { templates: Template[]; reviewers: Reviewer[]; tasks: Task[] };

const uid = () => Math.random().toString(36).slice(2, 9);

const excelColor = (value?: string) => {
  if (!value) return undefined;
  const normalized = value.replace(/^FF/i, '');
  return /^[0-9a-f]{6}$/i.test(normalized) ? `#${normalized}` : undefined;
};

function parseSheetLayout(workbook: XLSX.WorkBook) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet?.['!ref']) return null;
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const rowCount = Math.min(range.e.r + 1, 200);
  const colCount = Math.min(range.e.c + 1, 40);
  const merges = (sheet['!merges'] || []).filter(
    (merge) => merge.s.r < rowCount && merge.s.c < colCount,
  );
  const covered = new Set<string>();
  const mergeAt = new Map<string, XLSX.Range>();
  for (const merge of merges) {
    mergeAt.set(`${merge.s.r}:${merge.s.c}`, merge);
    for (let row = merge.s.r; row <= merge.e.r; row += 1) {
      for (let col = merge.s.c; col <= merge.e.c; col += 1) {
        if (row !== merge.s.r || col !== merge.s.c)
          covered.add(`${row}:${col}`);
      }
    }
  }
  const textAt = (row: number, col: number) => {
    const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
    return String(cell?.w ?? cell?.v ?? '').trim();
  };
  const scoreHeaderRow = Array.from({ length: rowCount }, (_, row) => row).find(
    (row) =>
      Array.from({ length: colCount }, (_, col) => textAt(row, col)).some(
        (text) => /^(分\s*值|满分|分数)$/.test(text),
      ) &&
      Array.from({ length: colCount }, (_, col) =>
        Number(textAt(row, col)),
      ).filter((value) => Number.isFinite(value) && value > 0).length >= 2,
  );
  if (scoreHeaderRow === undefined) return null;
  const scoreColumns = Array.from({ length: colCount }, (_, col) => ({
    col,
    max: Number(textAt(scoreHeaderRow, col)),
  })).filter(({ max }) => Number.isFinite(max) && max > 0);
  if (!scoreColumns.length) return null;
  const inputRow = Array.from(
    { length: rowCount - scoreHeaderRow - 1 },
    (_, index) => scoreHeaderRow + index + 1,
  ).find((row) =>
    Array.from({ length: Math.min(3, colCount) }, (_, col) =>
      textAt(row, col),
    ).some((text) => /^(合计|评分|得分|实际得分)$/.test(text)),
  );
  if (inputRow === undefined) return null;
  const criteria: Criterion[] = scoreColumns.map(({ col, max }, index) => {
    let criterionName = '';
    for (
      let row = scoreHeaderRow - 1;
      row >= 0 && row >= scoreHeaderRow - 4;
      row -= 1
    ) {
      const candidate = textAt(row, col);
      if (candidate && !/^考核评分要素$/.test(candidate)) {
        criterionName = candidate;
        break;
      }
    }
    return {
      id: `sheet-score-${index + 1}-${uid()}`,
      name: criterionName || `评分项 ${index + 1}`,
      dimension: '原表评分项',
      weight: max,
      max,
      description: '',
    };
  });
  const scoreByColumn = new Map(
    scoreColumns.map(({ col }, index) => [col, criteria[index]]),
  );
  let remarksCell: { row: number; col: number } | undefined;
  for (let row = 0; row <= scoreHeaderRow; row += 1) {
    for (let col = 0; col < colCount; col += 1) {
      if (/^备注$/.test(textAt(row, col))) {
        const headerMerge = mergeAt.get(`${row}:${col}`);
        remarksCell = {
          row: Math.max(scoreHeaderRow + 1, (headerMerge?.e.r ?? row) + 1),
          col,
        };
      }
    }
  }
  const cells: SheetCell[] = [];
  for (let row = 0; row < rowCount; row += 1) {
    for (let col = 0; col < colCount; col += 1) {
      if (covered.has(`${row}:${col}`)) continue;
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      const source = sheet[address] as
        | (XLSX.CellObject & { s?: Record<string, any> })
        | undefined;
      const merge = mergeAt.get(`${row}:${col}`);
      const rawStyle =
        source?.s && typeof source.s === 'object' ? source.s : undefined;
      const text = textAt(row, col);
      const cell: SheetCell = {
        row,
        col,
        rowSpan: merge ? merge.e.r - merge.s.r + 1 : 1,
        colSpan: merge ? merge.e.c - merge.s.c + 1 : 1,
        text,
        style: rawStyle
          ? {
              bold: Boolean(rawStyle.font?.bold),
              fontSize: Number(rawStyle.font?.sz) || undefined,
              align: rawStyle.alignment?.horizontal,
              vertical: rawStyle.alignment?.vertical,
              background: excelColor(rawStyle.fill?.fgColor?.rgb),
              color: excelColor(rawStyle.font?.color?.rgb),
              wrap: Boolean(rawStyle.alignment?.wrapText),
            }
          : undefined,
      };
      const scoreCriterion = scoreByColumn.get(col);
      if (row === remarksCell?.row && col === remarksCell.col) {
        cell.role = 'remarks';
      } else if (row === inputRow && scoreCriterion) {
        cell.role = 'score';
        cell.criterionId = scoreCriterion.id;
      } else if (/综合得分|总分/.test(text)) {
        const targetCol = Math.min(col + cell.colSpan, colCount - 1);
        const targetKey = `${row}:${targetCol}`;
        const targetMerge = mergeAt.get(targetKey);
        cells.push(cell);
        cells.push({
          row,
          col: targetCol,
          rowSpan: targetMerge ? targetMerge.e.r - targetMerge.s.r + 1 : 1,
          colSpan: targetMerge ? targetMerge.e.c - targetMerge.s.c + 1 : 1,
          text: '',
          role: 'total',
        });
        continue;
      } else if (/^综合评价/.test(text)) {
        const targetCol = Math.min(col + cell.colSpan, colCount - 1);
        const target = mergeAt.get(`${row}:${targetCol}`);
        cells.push(cell);
        cells.push({
          row,
          col: targetCol,
          rowSpan: target ? target.e.r - target.s.r + 1 : 1,
          colSpan: target ? target.e.c - target.s.c + 1 : 1,
          text: '',
          role: 'comment',
        });
        continue;
      } else if (/汇报人|被评价人|被评价对象/.test(text)) {
        cell.role = 'subject';
      } else if (/^评价人/.test(text)) {
        cell.role = 'reviewer';
      } else if (/考察.*阶段|考察期/.test(text)) {
        cell.role = 'stage';
      } else if (/评价者签名|评价人签名/.test(text)) {
        cell.role = 'signature';
      }
      cells.push(cell);
    }
  }
  const cellMap = new Map<string, SheetCell>();
  for (const cell of cells) {
    const key = `${cell.row}:${cell.col}`;
    if (cellMap.get(key)?.role && !cell.role) continue;
    cellMap.set(key, cell);
  }
  const uniqueCells = Array.from(cellMap.values());
  return {
    criteria,
    layout: {
      sheetName,
      rowCount,
      colCount,
      columnWidths: Array.from({ length: colCount }, (_, col) =>
        Math.max(
          58,
          Math.min(420, Number(sheet['!cols']?.[col]?.wch || 12) * 8),
        ),
      ),
      rowHeights: Array.from({ length: rowCount }, (_, row) =>
        Math.max(
          34,
          Math.min(420, Number(sheet['!rows']?.[row]?.hpt || 26) * 1.34),
        ),
      ),
      cells: uniqueCells,
    } satisfies SheetLayout,
  };
}
const today = () =>
  new Date().toLocaleDateString('zh-CN', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  });
const formatStageDate = (value?: string) => {
  if (!value) return '未设置';
  const [year, month, day] = value.split('-');
  return year && month && day
    ? `${year}年${Number(month)}月${Number(day)}日`
    : value;
};
const formatStageRange = (task: Task) => {
  const start = task.stageStartDate || task.stageDate;
  const end = task.stageEndDate || task.stageDate;
  if (!start || !end) return '未设置';
  return `${formatStageDate(start)}—${formatStageDate(end)}`;
};

const baseCriteria: Criterion[] = [
  {
    id: 'c1',
    name: '目标达成度',
    dimension: '工作绩效',
    weight: 30,
    max: 100,
    description: '按期、按质完成关键目标',
  },
  {
    id: 'c2',
    name: '专业能力',
    dimension: '能力素质',
    weight: 25,
    max: 100,
    description: '专业知识、问题分析与解决能力',
  },
  {
    id: 'c3',
    name: '协作意识',
    dimension: '团队协作',
    weight: 20,
    max: 100,
    description: '主动配合、信息共享、团队贡献',
  },
  {
    id: 'c4',
    name: '责任担当',
    dimension: '职业素养',
    weight: 15,
    max: 100,
    description: '对结果负责，主动承担挑战',
  },
  {
    id: 'c5',
    name: '创新改进',
    dimension: '发展潜力',
    weight: 10,
    max: 100,
    description: '持续优化方法并提出有效建议',
  },
];

const defaultTemplate: Template = {
  id: 'tpl-default',
  name: '季度综合评价表',
  description: '适用于员工季度绩效与能力综合评价',
  criteria: baseCriteria,
  updatedAt: '2026/08/28',
};

const stageReportTemplate: Template = {
  id: 'tpl-stage-report',
  name: '阶段汇报评价表',
  description:
    '用于试用期或考察期阶段汇报，评价人根据工作成果、关键能力和现场答辩进行百分制评分。',
  updatedAt: '2026/09/01',
  criteria: [
    {
      id: 'stage-work-result',
      name: '试用期工作成果及工作规划',
      dimension: '工作成果与规划',
      weight: 60,
      max: 60,
      description:
        '目标完成情况：工作指标完成情况、进度和质量；试用期工作成果：用数据或具体案例说明成果及解决的问题；工作规划：下一阶段的工作计划、任务、节点和重点安排。评分参考：好 44–60 分，中 21–43 分，差 1–20 分。',
    },
    {
      id: 'stage-key-ability',
      name: '关键能力',
      dimension: '能力素质',
      weight: 30,
      max: 30,
      description:
        '沟通协作：善于沟通、团结同事并协调内外部关系；学习能力：主动学习、自驱力强；执行力：按时按质按量完成目标；问题解决与推动能力：结果导向，能预见问题并推动解决。评分参考：好 26–30 分，中 16–25 分，差 1–15 分。',
    },
    {
      id: 'stage-defense',
      name: '现场答辩',
      dimension: '汇报答辩',
      weight: 10,
      max: 10,
      description:
        '准确性：回答直击要点；条理性：论点及论证逻辑清晰、结构合理、内容关联性强；答辩效果：思路敏捷、语言组织和归纳总结能力强、现场反应效果好。评分参考：好 9–10 分，中 6–8 分，差 1–5 分。',
    },
  ],
};

const seedReviewers: Reviewer[] = [
  {
    id: 'reviewer-wang',
    name: '王主管',
    department: '管理层',
    code: 'WANG-8264',
    active: true,
  },
  {
    id: 'reviewer-zhao',
    name: '赵同事',
    department: '员工代表',
    code: 'ZHAO-5179',
    active: true,
  },
];

const seed: AppData = {
  templates: [defaultTemplate, stageReportTemplate],
  reviewers: seedReviewers,
  tasks: [
    {
      id: 'task-q3',
      name: '2026 年第三季度综合评价',
      templateId: 'tpl-default',
      deadline: '2026-09-15',
      createdAt: '2026/08/28',
      assignments: [
        {
          id: 'a1',
          subject: '李明',
          department: '产品中心',
          reviewer: '王主管',
          reviewerId: 'reviewer-wang',
          status: 'submitted',
          scores: { c1: 94, c2: 92, c3: 90, c4: 91, c5: 88 },
          comment: '目标意识强，推动跨部门协作很有成效。',
          submittedAt: '2026/08/30',
        },
        {
          id: 'a2',
          subject: '李明',
          department: '产品中心',
          reviewer: '赵同事',
          reviewerId: 'reviewer-zhao',
          status: 'submitted',
          scores: { c1: 92, c2: 90, c3: 94, c4: 88, c5: 86 },
          comment: '沟通清晰，愿意主动补位。',
          submittedAt: '2026/08/31',
        },
        {
          id: 'a3',
          subject: '周宁',
          department: '运营中心',
          reviewer: '王主管',
          reviewerId: 'reviewer-wang',
          status: 'submitted',
          scores: { c1: 90, c2: 88, c3: 86, c4: 91, c5: 82 },
          comment: '执行稳定，可进一步沉淀方法。',
          submittedAt: '2026/08/31',
        },
        {
          id: 'a4',
          subject: '周宁',
          department: '运营中心',
          reviewer: '赵同事',
          reviewerId: 'reviewer-zhao',
          status: 'pending',
          scores: {},
          comment: '',
        },
        {
          id: 'a5',
          subject: '陈嘉',
          department: '技术中心',
          reviewer: '王主管',
          reviewerId: 'reviewer-wang',
          status: 'submitted',
          scores: { c1: 86, c2: 94, c3: 82, c4: 88, c5: 91 },
          comment: '专业能力突出，建议增加业务沟通。',
          submittedAt: '2026/08/29',
        },
        {
          id: 'a6',
          subject: '陈嘉',
          department: '技术中心',
          reviewer: '赵同事',
          reviewerId: 'reviewer-zhao',
          status: 'pending',
          scores: {},
          comment: '',
        },
      ],
    },
  ],
};

function normalizeData(raw: AppData): AppData {
  const knownReviewers =
    Array.isArray(raw.reviewers) && raw.reviewers.length
      ? raw.reviewers
      : seedReviewers;
  const templates: Template[] = raw.templates.map((template) => ({
    ...template,
    version: template.version || 1,
  }));
  if (!templates.some((template) => template.id === defaultTemplate.id))
    templates.unshift(defaultTemplate);
  if (!templates.some((template) => template.id === stageReportTemplate.id))
    templates.push(stageReportTemplate);
  const tasks = raw.tasks.map((task) => ({
    ...task,
    stageStartDate: task.stageStartDate || task.stageDate,
    stageEndDate: task.stageEndDate || task.stageDate,
    assignments: task.assignments.map((assignment) => ({
      ...assignment,
      reviewerId:
        assignment.reviewerId ||
        knownReviewers.find((r) => r.name === assignment.reviewer)?.id,
    })),
  }));
  return { ...raw, templates, reviewers: knownReviewers, tasks };
}

function weightedScore(a: Assignment, tpl?: Template) {
  if (!tpl || a.status !== 'submitted') return 0;
  const totalWeight = tpl.criteria.reduce((s, c) => s + c.weight, 0) || 1;
  return (
    (tpl.criteria.reduce(
      (s, c) => s + ((a.scores[c.id] || 0) / c.max) * c.weight,
      0,
    ) /
      totalWeight) *
    100
  );
}

export default function EvaluationApp() {
  const [data, setData] = useState<AppData>(seed);
  const [ready, setReady] = useState(false);
  const [storageMode, setStorageMode] = useState<
    'connecting' | 'server' | 'browser'
  >('connecting');
  const skipNextPersist = useRef(false);
  const dataRef = useRef(data);
  dataRef.current = data;
  const [view, setView] = useState<View>('dashboard');
  const [reviewerOnlyMode, setReviewerOnlyMode] = useState(false);
  const [mobileNav, setMobileNav] = useState(false);
  const [templateOpen, setTemplateOpen] = useState(false);
  const [editingTemplate, setEditingTemplate] = useState<Template | null>(null);
  const [taskOpen, setTaskOpen] = useState(false);
  const [reviewOpen, setReviewOpen] = useState(false);
  const [activeAssignment, setActiveAssignment] = useState<{
    taskId: string;
    assignmentId: string;
  } | null>(null);
  const [notice, setNotice] = useState('');

  useEffect(() => {
    const load = async () => {
      try {
        const response = await fetch('/api/data', { cache: 'no-store' });
        if (!response.ok) throw new Error('server unavailable');
        const payload = (await response.json()) as { data: AppData | null };
        let serverData = payload.data;
        if (!serverData) {
          try {
            const localData = localStorage.getItem('pingjian-data-v1');
            serverData = localData ? (JSON.parse(localData) as AppData) : seed;
          } catch {
            serverData = seed;
          }
        }
        const migrated = normalizeData(serverData);
        skipNextPersist.current = true;
        setData(migrated);
        setStorageMode('server');
        if (
          !payload.data ||
          JSON.stringify(migrated) !== JSON.stringify(serverData)
        ) {
          await fetch('/api/data', {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ data: migrated }),
          });
        }
      } catch {
        try {
          const saved = localStorage.getItem('pingjian-data-v1');
          if (saved) {
            const parsed = JSON.parse(saved) as AppData;
            setData(normalizeData(parsed));
          }
        } catch {
          /* use seed */
        }
        setStorageMode('browser');
      } finally {
        setReady(true);
      }
    };
    void load();
  }, []);
  useEffect(() => {
    if (!ready) return;
    const params = new URLSearchParams(window.location.search);
    if (params.get('portal') === 'reviewer') {
      setReviewerOnlyMode(true);
      setView('reviewer');
    }
    const reviewId = params.get('review');
    if (!reviewId) return;
    const task = data.tasks.find((t) =>
      t.assignments.some((a) => a.id === reviewId),
    );
    if (task) {
      setReviewerOnlyMode(true);
      setView('reviewer');
    }
  }, [ready]);
  useEffect(() => {
    if (!ready) return;
    if (skipNextPersist.current) {
      skipNextPersist.current = false;
      return;
    }
    if (storageMode === 'server') {
      const timer = setTimeout(() => {
        void fetch('/api/data', {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ data }),
        })
          .then((response) => {
            if (!response.ok) throw new Error('save failed');
          })
          .catch(() => setNotice('服务器保存失败，请稍后重试'));
      }, 250);
      return () => clearTimeout(timer);
    }
    if (storageMode === 'browser')
      localStorage.setItem('pingjian-data-v1', JSON.stringify(data));
  }, [data, ready, storageMode]);
  useEffect(() => {
    if (!ready || storageMode !== 'server') return;
    const timer = setInterval(() => {
      void fetch('/api/data', { cache: 'no-store' })
        .then(async (response) => {
          if (!response.ok) throw new Error('sync failed');
          return (await response.json()) as { data: AppData | null };
        })
        .then((payload) => {
          if (!payload.data) return;
          const normalized = normalizeData(payload.data);
          if (JSON.stringify(normalized) !== JSON.stringify(dataRef.current)) {
            skipNextPersist.current = true;
            setData(normalized);
          }
        })
        .catch(() => undefined);
    }, 5000);
    return () => clearInterval(timer);
  }, [ready, storageMode]);
  useEffect(() => {
    if (!notice) return;
    const timer = setTimeout(() => setNotice(''), 2600);
    return () => clearTimeout(timer);
  }, [notice]);

  const stats = useMemo(() => {
    const all = data.tasks.flatMap((t) =>
      t.assignments.map((a) => ({
        a,
        tpl: data.templates.find((x) => x.id === t.templateId),
      })),
    );
    const submitted = all.filter((x) => x.a.status === 'submitted');
    const avg = submitted.length
      ? submitted.reduce((s, x) => s + weightedScore(x.a, x.tpl), 0) /
        submitted.length
      : 0;
    return {
      active: data.tasks.length,
      total: all.length,
      submitted: submitted.length,
      rate: all.length ? (submitted.length / all.length) * 100 : 0,
      avg,
    };
  }, [data]);

  const subjectResults = useMemo(() => {
    const map = new Map<
      string,
      {
        subject: string;
        department: string;
        scores: number[];
        total: number;
        submitted: number;
      }
    >();
    data.tasks.forEach((t) => {
      const tpl = data.templates.find((x) => x.id === t.templateId);
      t.assignments.forEach((a) => {
        const key = `${a.subject}-${a.department}`;
        const row = map.get(key) || {
          subject: a.subject,
          department: a.department,
          scores: [],
          total: 0,
          submitted: 0,
        };
        row.total++;
        if (a.status === 'submitted') {
          row.submitted++;
          row.scores.push(weightedScore(a, tpl));
        }
        map.set(key, row);
      });
    });
    return [...map.values()]
      .map((x) => ({
        ...x,
        score: x.scores.length
          ? x.scores.reduce((a, b) => a + b, 0) / x.scores.length
          : 0,
      }))
      .sort((a, b) => b.score - a.score);
  }, [data]);

  const nav = [
    { id: 'dashboard' as View, label: '总览', icon: LayoutDashboard },
    { id: 'templates' as View, label: '评价表模板', icon: FileSpreadsheet },
    { id: 'reviewers' as View, label: '评价人管理', icon: UserPlus },
    { id: 'tasks' as View, label: '对象与任务', icon: Users },
    { id: 'results' as View, label: '结果汇总', icon: BarChart3 },
    { id: 'reviewer' as View, label: '评价填写入口', icon: PencilLine },
  ];

  const navigate = (next: View) => {
    setView(next);
    setMobileNav(false);
  };
  const openAssignment = (taskId: string, assignmentId: string) => {
    setActiveAssignment({ taskId, assignmentId });
    setReviewOpen(true);
  };

  if (view === 'reviewer') {
    return (
      <main className="min-h-screen bg-[#f4f7f5] text-[#183028]">
        {notice && (
          <div className="fixed right-5 top-5 z-[80] flex items-center gap-2 rounded-xl bg-[#123c32] px-4 py-3 text-sm text-white shadow-xl">
            <Check className="size-4 text-[#d9ff70]" />
            {notice}
          </div>
        )}
        <header className="border-b border-[#dce5e1] bg-[#123c32] text-white">
          <div className="mx-auto flex h-[72px] max-w-5xl items-center justify-between px-5">
            <div className="flex items-center gap-3">
              <span className="grid size-9 place-items-center rounded-xl bg-[#d9ff70] text-[#123c32]">
                <ClipboardCheck className="size-5" />
              </span>
              <div>
                <p className="font-semibold tracking-tight">评鉴</p>
                <p className="text-[11px] text-white/55">评价填写入口</p>
              </div>
            </div>
            {!reviewerOnlyMode && (
              <Button
                variant="outline"
                className="border-white/20 bg-white/10 text-white hover:bg-white/15 hover:text-white"
                onClick={() => setView('dashboard')}
              >
                返回管理后台
              </Button>
            )}
          </div>
        </header>
        <div className="mx-auto max-w-5xl p-5 md:p-8">
          <ReviewerPortal data={data} onOpen={openAssignment} />
        </div>
        <ReviewDialog
          open={reviewOpen}
          onOpenChange={setReviewOpen}
          active={activeAssignment}
          data={data}
          setData={setData}
          setNotice={setNotice}
        />
      </main>
    );
  }

  return (
    <main className="min-h-screen bg-[#f4f7f5] text-[#183028]">
      {notice && (
        <div className="fixed right-5 top-5 z-[80] flex items-center gap-2 rounded-xl bg-[#123c32] px-4 py-3 text-sm text-white shadow-xl">
          <Check className="size-4 text-[#d9ff70]" />
          {notice}
        </div>
      )}
      <aside
        className={`fixed inset-y-0 left-0 z-50 w-[236px] border-r border-white/10 bg-[#123c32] px-4 py-6 text-white transition-transform lg:translate-x-0 ${mobileNav ? 'translate-x-0' : '-translate-x-full'}`}
      >
        <div className="flex items-center justify-between px-2">
          <div className="flex items-center gap-3">
            <span className="grid size-9 place-items-center rounded-xl bg-[#d9ff70] text-[#123c32]">
              <ClipboardCheck className="size-5" />
            </span>
            <div>
              <p className="font-semibold tracking-tight">评鉴</p>
              <p className="text-[11px] text-white/55">企业内网评价工作台</p>
            </div>
          </div>
          <button
            className="lg:hidden"
            onClick={() => setMobileNav(false)}
            aria-label="关闭菜单"
          >
            <X className="size-5" />
          </button>
        </div>
        <nav className="mt-9 space-y-1 text-sm">
          {nav.map((item) => (
            <button
              key={item.id}
              onClick={() => navigate(item.id)}
              className={`flex w-full items-center gap-3 rounded-xl px-3 py-2.5 text-left transition ${view === item.id ? 'bg-white/13 font-medium text-white' : 'text-white/62 hover:bg-white/8 hover:text-white'}`}
            >
              <item.icon className="size-4" />
              {item.label}
            </button>
          ))}
        </nav>
        <div className="absolute inset-x-4 bottom-5 rounded-2xl bg-white/8 p-3 text-xs text-white/60">
          <div className="mb-2 flex items-center gap-2 text-white">
            <Settings className="size-4" />
            {storageMode === 'server'
              ? '服务器共享模式'
              : storageMode === 'connecting'
                ? '正在连接服务器'
                : '浏览器备用模式'}
          </div>
          {storageMode === 'server'
            ? '局域网内所有电脑共享数据。'
            : storageMode === 'connecting'
              ? '正在读取统一评价数据。'
              : '服务器不可用，数据临时保存在本机。'}
        </div>
      </aside>
      {mobileNav && (
        <button
          className="fixed inset-0 z-40 bg-black/25 lg:hidden"
          onClick={() => setMobileNav(false)}
          aria-label="关闭菜单遮罩"
        />
      )}

      <section className="lg:pl-[236px]">
        <header className="sticky top-0 z-30 flex h-[72px] items-center justify-between border-b border-[#dce5e1] bg-white/88 px-5 backdrop-blur md:px-8">
          <div className="flex items-center gap-3">
            <button
              className="lg:hidden"
              onClick={() => setMobileNav(true)}
              aria-label="打开菜单"
            >
              <Menu className="size-5" />
            </button>
            <div>
              <p className="text-[11px] text-[#718078]">2026 年度 · 评价管理</p>
              <h1 className="font-semibold tracking-tight">
                {nav.find((x) => x.id === view)?.label}
              </h1>
            </div>
          </div>
          <div className="flex gap-2">
            <Button
              variant="outline"
              className="hidden h-10 rounded-xl sm:flex"
              onClick={() => navigate('reviewer')}
            >
              <Eye />
              评价入口
            </Button>
            <Button
              className="h-10 rounded-xl bg-[#123c32] px-4 hover:bg-[#1c5144]"
              onClick={() => setTaskOpen(true)}
            >
              <Plus />
              新建任务
            </Button>
          </div>
        </header>

        <div className="mx-auto max-w-[1240px] p-5 md:p-8">
          {view === 'dashboard' && (
            <Dashboard
              stats={stats}
              results={subjectResults}
              onNew={() => setTaskOpen(true)}
              onViewResults={() => navigate('results')}
            />
          )}
          {view === 'templates' && (
            <Templates
              data={data}
              setData={setData}
              setNotice={setNotice}
              onNew={() => {
                setEditingTemplate(null);
                setTemplateOpen(true);
              }}
              onEdit={(template) => {
                setEditingTemplate(template);
                setTemplateOpen(true);
              }}
            />
          )}
          {view === 'reviewers' && (
            <ReviewerManager
              data={data}
              setData={setData}
              setNotice={setNotice}
            />
          )}
          {view === 'tasks' && (
            <Tasks
              data={data}
              setData={setData}
              onNew={() => setTaskOpen(true)}
              onOpenAssignment={openAssignment}
              setNotice={setNotice}
            />
          )}
          {view === 'results' && (
            <Results
              data={data}
              results={subjectResults}
              setNotice={setNotice}
            />
          )}
        </div>
      </section>

      <TemplateDialog
        open={templateOpen}
        onOpenChange={(nextOpen) => {
          setTemplateOpen(nextOpen);
          if (!nextOpen) setEditingTemplate(null);
        }}
        setData={setData}
        setNotice={setNotice}
        template={editingTemplate}
      />
      <TaskDialog
        open={taskOpen}
        onOpenChange={setTaskOpen}
        data={data}
        setData={setData}
        setNotice={setNotice}
      />
      <ReviewDialog
        open={reviewOpen}
        onOpenChange={setReviewOpen}
        active={activeAssignment}
        data={data}
        setData={setData}
        setNotice={setNotice}
      />
    </main>
  );
}

function Dashboard({
  stats,
  results,
  onNew,
  onViewResults,
}: {
  stats: {
    active: number;
    total: number;
    submitted: number;
    rate: number;
    avg: number;
  };
  results: ReturnType<any>;
  onNew: () => void;
  onViewResults: () => void;
}) {
  return (
    <div className="space-y-6">
      <section className="flex flex-col justify-between gap-4 rounded-[22px] bg-[#123c32] p-6 text-white md:flex-row md:items-center md:p-7">
        <div>
          <p className="text-sm text-white/55">
            把线下表格变成可追踪的在线评价
          </p>
          <h2 className="mt-1 text-2xl font-semibold tracking-tight">
            从模板到结果，一处完成
          </h2>
          <p className="mt-2 max-w-xl text-sm leading-6 text-white/62">
            上传现有表格或自定义指标，分配给评价人，系统会自动计算每个对象的综合得分。
          </p>
        </div>
        <Button
          className="h-11 shrink-0 rounded-xl bg-[#d9ff70] px-5 text-[#123c32] hover:bg-[#c9ef62]"
          onClick={onNew}
        >
          <Plus />
          创建评价任务
        </Button>
      </section>
      <div className="grid gap-4 md:grid-cols-3">
        <Stat
          title="进行中的任务"
          value={`${stats.active}`}
          hint={`共 ${stats.total} 份评价，已回收 ${stats.submitted} 份`}
          icon={ClipboardCheck}
        />
        <Stat
          title="整体完成率"
          value={`${stats.rate.toFixed(1)}%`}
          hint="评价人提交后实时更新"
          icon={Sparkles}
          progress={stats.rate}
        />
        <Stat
          title="当前综合均分"
          value={stats.avg.toFixed(1)}
          hint="按指标权重自动折算"
          icon={BarChart3}
          accent
        />
      </div>
      <article className="overflow-hidden rounded-2xl border border-[#dce5e1] bg-white shadow-[0_8px_30px_rgba(23,48,40,.04)]">
        <div className="flex items-center justify-between border-b border-[#e4ebe8] px-5 py-4">
          <div>
            <h2 className="font-semibold">被评价对象</h2>
            <p className="mt-0.5 text-xs text-[#718078]">按综合得分排序</p>
          </div>
          <button
            onClick={onViewResults}
            className="text-sm font-medium text-[#1c6b56]"
          >
            查看全部 →
          </button>
        </div>
        <ResultTable results={results} />
      </article>
    </div>
  );
}

function Stat({
  title,
  value,
  hint,
  icon: Icon,
  progress,
  accent,
}: {
  title: string;
  value: string;
  hint: string;
  icon: any;
  progress?: number;
  accent?: boolean;
}) {
  return (
    <article
      className={`rounded-2xl p-5 shadow-[0_8px_30px_rgba(23,48,40,.04)] ${accent ? 'bg-[#d9ff70]' : 'border border-[#dce5e1] bg-white'}`}
    >
      <div className="flex items-start justify-between">
        <p
          className={`text-sm ${accent ? 'text-[#486133]' : 'text-[#718078]'}`}
        >
          {title}
        </p>
        <span
          className={`rounded-lg p-2 ${accent ? 'bg-white/55' : 'bg-[#eaf5f0] text-[#1c6b56]'}`}
        >
          <Icon className="size-4" />
        </span>
      </div>
      <p className="mt-3 text-3xl font-semibold tracking-tight">{value}</p>
      {progress !== undefined && (
        <div className="mt-3 h-1.5 rounded-full bg-[#edf1ef]">
          <div
            className="h-full rounded-full bg-[#b7db45]"
            style={{ width: `${progress}%` }}
          />
        </div>
      )}
      <p className="mt-2 text-xs text-[#64736b]">{hint}</p>
    </article>
  );
}

function Templates({
  data,
  setData,
  setNotice,
  onNew,
  onEdit,
}: {
  data: AppData;
  setData: React.Dispatch<React.SetStateAction<AppData>>;
  setNotice: (s: string) => void;
  onNew: () => void;
  onEdit: (template: Template) => void;
}) {
  return (
    <div className="space-y-5">
      <PageIntro
        title="评价表模板"
        text="部署后可直接上传 Excel：系统自动保存合并单元格、行列尺寸和评分位置，并按原表结构生成在线表。"
        action={
          <Button onClick={onNew} className="h-10 rounded-xl bg-[#123c32]">
            <Plus />
            新建或导入模板
          </Button>
        }
      />
      <div className="grid gap-4 lg:grid-cols-2">
        {data.templates.map((t) => (
          <article
            key={t.id}
            className="rounded-2xl border border-[#dce5e1] bg-white p-5 shadow-[0_8px_30px_rgba(23,48,40,.04)]"
          >
            <div className="flex items-start justify-between">
              <span className="grid size-10 place-items-center rounded-xl bg-[#eaf5f0] text-[#1c6b56]">
                <FileSpreadsheet className="size-5" />
              </span>
              <div className="flex items-center gap-1">
                <button
                  onClick={() => onEdit(t)}
                  className="rounded-lg p-2 text-[#527067] hover:bg-[#eef5f2] hover:text-[#145b48]"
                  aria-label={`复制编辑 ${t.name}`}
                  title="复制为新版本编辑"
                >
                  <PencilLine className="size-4" />
                </button>
                <button
                  onClick={() => {
                    if (
                      t.id === defaultTemplate.id ||
                      t.id === stageReportTemplate.id
                    ) {
                      setNotice('内置模板不能删除');
                      return;
                    }
                    if (data.tasks.some((task) => task.templateId === t.id)) {
                      setNotice('该模板正在被评价任务使用，不能删除');
                      return;
                    }
                    setData((d) => ({
                      ...d,
                      templates: d.templates.filter((x) => x.id !== t.id),
                    }));
                  }}
                  className="rounded-lg p-2 text-[#8a9690] hover:bg-red-50 hover:text-red-600"
                  aria-label="删除模板"
                >
                  <Trash2 className="size-4" />
                </button>
              </div>
            </div>
            <div className="mt-4 flex items-center gap-2">
              <h3 className="font-semibold">{t.name}</h3>
              <span className="rounded-full bg-[#f0f2f1] px-2 py-0.5 text-[11px] text-[#64736b]">
                v{t.version || 1}
              </span>
              {t.sheetLayout && (
                <span className="rounded-full bg-[#e4f4ec] px-2 py-0.5 text-[11px] font-medium text-[#17604c]">
                  原表布局
                </span>
              )}
            </div>
            <p className="mt-1 min-h-10 text-sm leading-5 text-[#718078]">
              {t.description}
            </p>
            <div className="mt-4 flex gap-4 border-t border-[#edf1ef] pt-4 text-xs text-[#718078]">
              <span>
                <b className="text-[#183028]">{t.criteria.length}</b> 个指标
              </span>
              <span>
                权重合计{' '}
                <b className="text-[#183028]">
                  {t.criteria.reduce((s, c) => s + c.weight, 0)}%
                </b>
              </span>
              <span className="ml-auto">更新于 {t.updatedAt}</span>
            </div>
            <details className="mt-4 rounded-xl bg-[#f6f8f7] p-3">
              <summary className="cursor-pointer text-sm font-medium">
                查看指标明细
              </summary>
              <div className="mt-3 space-y-2">
                {t.criteria.map((c) => (
                  <div
                    key={c.id}
                    className="flex items-center justify-between text-xs"
                  >
                    <span>
                      {c.dimension} · {c.name}
                    </span>
                    <span className="text-[#718078]">
                      {c.weight}% / {c.max}分
                    </span>
                  </div>
                ))}
              </div>
            </details>
          </article>
        ))}
      </div>
    </div>
  );
}

function ReviewerManager({
  data,
  setData,
  setNotice,
}: {
  data: AppData;
  setData: React.Dispatch<React.SetStateAction<AppData>>;
  setNotice: (value: string) => void;
}) {
  const [name, setName] = useState('');
  const [department, setDepartment] = useState('');
  const [code, setCode] = useState('');
  const makeCode = () =>
    `PJ-${Math.random().toString(36).slice(2, 8).toUpperCase()}`;
  const addReviewer = () => {
    const normalizedCode = code.trim().toUpperCase();
    if (!name.trim() || !normalizedCode) {
      setNotice('请填写评价人姓名和登录码');
      return;
    }
    if (data.reviewers.some((r) => r.code.toUpperCase() === normalizedCode)) {
      setNotice('该登录码已被使用，请更换');
      return;
    }
    setData((current) => ({
      ...current,
      reviewers: [
        ...current.reviewers,
        {
          id: uid(),
          name: name.trim(),
          department: department.trim() || '未分组',
          code: normalizedCode,
          active: true,
        },
      ],
    }));
    setName('');
    setDepartment('');
    setCode('');
    setNotice('评价人已添加');
  };
  const removeReviewer = (reviewer: Reviewer) => {
    if (
      data.tasks.some((task) =>
        task.assignments.some((a) => a.reviewerId === reviewer.id),
      )
    ) {
      setNotice('该评价人已有分配记录，不能删除，可将其停用');
      return;
    }
    setData((current) => ({
      ...current,
      reviewers: current.reviewers.filter((r) => r.id !== reviewer.id),
    }));
  };
  return (
    <div className="space-y-5">
      <PageIntro
        title="评价人管理"
        text="提前维护评价人信息和唯一登录码。评价人通过登录码进入自己的待评价清单。"
      />
      <article className="rounded-2xl border border-[#dce5e1] bg-white p-5">
        <div className="grid gap-3 md:grid-cols-[1fr_1fr_1fr_auto]">
          <label className="text-sm font-medium">
            姓名
            <Input
              className="mt-1 h-10"
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="评价人姓名"
            />
          </label>
          <label className="text-sm font-medium">
            部门 / 身份
            <Input
              className="mt-1 h-10"
              value={department}
              onChange={(e) => setDepartment(e.target.value)}
              placeholder="例如：管理层"
            />
          </label>
          <label className="text-sm font-medium">
            差异化登录码
            <div className="mt-1 flex gap-2">
              <Input
                className="h-10 uppercase"
                value={code}
                onChange={(e) => setCode(e.target.value)}
                placeholder="例如：PJ-A8K2M6"
              />
              <Button
                type="button"
                variant="outline"
                className="h-10"
                onClick={() => setCode(makeCode())}
              >
                生成
              </Button>
            </div>
          </label>
          <Button className="h-10 self-end bg-[#123c32]" onClick={addReviewer}>
            <UserPlus />
            添加评价人
          </Button>
        </div>
      </article>
      <div className="grid gap-4 lg:grid-cols-2">
        {data.reviewers.map((reviewer) => {
          const assigned = data.tasks.reduce(
            (sum, task) =>
              sum +
              task.assignments.filter(
                (a) =>
                  a.reviewerId === reviewer.id ||
                  (!a.reviewerId && a.reviewer === reviewer.name),
              ).length,
            0,
          );
          return (
            <article
              key={reviewer.id}
              className={`rounded-2xl border bg-white p-5 ${reviewer.active ? 'border-[#dce5e1]' : 'border-[#e6e8e7] opacity-65'}`}
            >
              <div className="flex items-start gap-3">
                <span className="grid size-10 place-items-center rounded-full bg-[#eaf5f0] font-semibold text-[#1c6b56]">
                  {reviewer.name[0]}
                </span>
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-2">
                    <h3 className="font-semibold">{reviewer.name}</h3>
                    <span
                      className={`rounded-full px-2 py-0.5 text-[11px] ${reviewer.active ? 'bg-[#eaf5f0] text-[#1c6b56]' : 'bg-[#edf1ef] text-[#718078]'}`}
                    >
                      {reviewer.active ? '启用' : '停用'}
                    </span>
                  </div>
                  <p className="mt-1 text-xs text-[#718078]">
                    {reviewer.department} · 已分配 {assigned} 份评价
                  </p>
                </div>
                <button
                  onClick={() => removeReviewer(reviewer)}
                  className="rounded-lg p-2 text-[#8a9690] hover:bg-red-50 hover:text-red-600"
                >
                  <Trash2 className="size-4" />
                </button>
              </div>
              <div className="mt-4 flex items-center gap-2 rounded-xl bg-[#f5f8f6] p-3">
                <KeyRound className="size-4 text-[#1c6b56]" />
                <code className="flex-1 text-sm font-semibold tracking-wider">
                  {reviewer.code}
                </code>
                <Button
                  size="sm"
                  variant="outline"
                  onClick={() => {
                    navigator.clipboard?.writeText(reviewer.code);
                    setNotice('登录码已复制');
                  }}
                >
                  <Copy />
                  复制
                </Button>
              </div>
              <div className="mt-3 flex justify-end">
                <Button
                  size="sm"
                  variant="ghost"
                  onClick={() =>
                    setData((current) => ({
                      ...current,
                      reviewers: current.reviewers.map((r) =>
                        r.id === reviewer.id ? { ...r, active: !r.active } : r,
                      ),
                    }))
                  }
                >
                  {reviewer.active ? '停用评价人' : '重新启用'}
                </Button>
              </div>
            </article>
          );
        })}
      </div>
    </div>
  );
}

function Tasks({
  data,
  setData,
  onNew,
  onOpenAssignment,
  setNotice,
}: {
  data: AppData;
  setData: React.Dispatch<React.SetStateAction<AppData>>;
  onNew: () => void;
  onOpenAssignment: (t: string, a: string) => void;
  setNotice: (s: string) => void;
}) {
  return (
    <div className="space-y-5">
      <PageIntro
        title="评价任务"
        text="为指定对象和评价人建立任务，跟踪回收状态，并可直接打开评价表。"
        action={
          <Button onClick={onNew} className="h-10 rounded-xl bg-[#123c32]">
            <Plus />
            新建任务
          </Button>
        }
      />
      {data.tasks.map((task) => {
        const done = task.assignments.filter(
          (a) => a.status === 'submitted',
        ).length;
        const tpl = data.templates.find((t) => t.id === task.templateId);
        return (
          <article
            key={task.id}
            className="overflow-hidden rounded-2xl border border-[#dce5e1] bg-white"
          >
            <div className="flex flex-col gap-3 border-b border-[#e5ece8] p-5 md:flex-row md:items-center md:justify-between">
              <div>
                <div className="flex items-center gap-2">
                  <h3 className="font-semibold">{task.name}</h3>
                  <span className="rounded-full bg-[#eaf5f0] px-2 py-0.5 text-[11px] text-[#1c6b56]">
                    进行中
                  </span>
                </div>
                <p className="mt-1 text-xs text-[#718078]">
                  {tpl?.name} · 截止 {task.deadline}
                </p>
                {tpl?.id === 'tpl-stage-report' && (
                  <div className="mt-3 flex flex-wrap items-center gap-2 text-xs font-medium text-[#516159]">
                    <span>考察阶段</span>
                    <Input
                      aria-label="考察阶段开始日期"
                      type="date"
                      value={task.stageStartDate || task.stageDate || ''}
                      onChange={(e) =>
                        setData((d) => ({
                          ...d,
                          tasks: d.tasks.map((item) =>
                            item.id === task.id
                              ? { ...item, stageStartDate: e.target.value }
                              : item,
                          ),
                        }))
                      }
                      className="h-8 w-40 bg-white text-xs"
                    />
                    <span>至</span>
                    <Input
                      aria-label="考察阶段结束日期"
                      type="date"
                      min={task.stageStartDate || task.stageDate || undefined}
                      value={task.stageEndDate || task.stageDate || ''}
                      onChange={(e) =>
                        setData((d) => ({
                          ...d,
                          tasks: d.tasks.map((item) =>
                            item.id === task.id
                              ? { ...item, stageEndDate: e.target.value }
                              : item,
                          ),
                        }))
                      }
                      className="h-8 w-40 bg-white text-xs"
                    />
                  </div>
                )}
              </div>
              <div className="flex items-center gap-3">
                <div className="text-right text-xs text-[#718078]">
                  <b className="text-base text-[#183028]">{done}</b> /{' '}
                  {task.assignments.length}
                  <br />
                  已提交
                </div>
                <button
                  onClick={() =>
                    setData((d) => ({
                      ...d,
                      tasks: d.tasks.filter((x) => x.id !== task.id),
                    }))
                  }
                  className="rounded-lg p-2 text-[#8a9690] hover:bg-red-50 hover:text-red-600"
                >
                  <Trash2 className="size-4" />
                </button>
              </div>
            </div>
            <div className="divide-y divide-[#edf1ef]">
              {task.assignments.map((a) => (
                <div
                  key={a.id}
                  className="grid gap-3 px-5 py-3 text-sm md:grid-cols-[1.2fr_1fr_1fr_auto] md:items-center"
                >
                  <div>
                    <b>{a.subject}</b>
                    <span className="ml-2 text-xs text-[#839089]">
                      {a.department}
                    </span>
                  </div>
                  <span className="text-[#64736b]">评价人：{a.reviewer}</span>
                  <span
                    className={
                      a.status === 'submitted'
                        ? 'text-[#1c765b]'
                        : 'text-[#a2781c]'
                    }
                  >
                    {a.status === 'submitted'
                      ? `已提交 · ${weightedScore(a, tpl).toFixed(1)}分`
                      : '待填写'}
                  </span>
                  <div className="flex gap-2">
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={() => {
                        navigator.clipboard?.writeText(
                          `${location.origin}/?portal=reviewer`,
                        );
                        setNotice('评价人登录入口已复制');
                      }}
                    >
                      <Copy />
                      登录入口
                    </Button>
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={() => onOpenAssignment(task.id, a.id)}
                    >
                      {a.status === 'submitted' ? <Eye /> : <PencilLine />}
                      {a.status === 'submitted' ? '查看' : '填写'}
                    </Button>
                  </div>
                </div>
              ))}
            </div>
          </article>
        );
      })}
    </div>
  );
}

function Results({
  data,
  results,
  setNotice,
}: {
  data: AppData;
  results: any[];
  setNotice: (s: string) => void;
}) {
  const exportCsv = () => {
    const rows = [
      ['被评价对象', '部门', '已提交', '应提交', '综合得分'],
      ...results.map((r) => [
        r.subject,
        r.department,
        r.submitted,
        r.total,
        r.score.toFixed(1),
      ]),
    ];
    const csv =
      '\ufeff' +
      rows
        .map((r) =>
          r.map((v) => `"${String(v).replaceAll('"', '""')}"`).join(','),
        )
        .join('\n');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(
      new Blob([csv], { type: 'text/csv;charset=utf-8' }),
    );
    a.download = `评价汇总-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(a.href);
    setNotice('汇总结果已导出');
  };
  return (
    <div className="space-y-5">
      <PageIntro
        title="综合评价结果"
        text="相同对象的多份评价自动合并，按模板权重折算为百分制综合分。"
        action={
          <Button
            variant="outline"
            className="h-10 rounded-xl"
            onClick={exportCsv}
          >
            <Download />
            导出汇总 CSV
          </Button>
        }
      />
      <div className="grid gap-4 md:grid-cols-3">
        <Stat
          title="被评价对象"
          value={`${results.length}`}
          hint="自动按姓名与部门合并"
          icon={Users}
        />
        <Stat
          title="已收到评价"
          value={`${data.tasks.flatMap((t) => t.assignments).filter((a) => a.status === 'submitted').length}`}
          hint="提交后自动进入汇总"
          icon={ClipboardCheck}
        />
        <Stat
          title="最高综合得分"
          value={results[0]?.score ? results[0].score.toFixed(1) : '—'}
          hint={results[0]?.subject || '暂无结果'}
          icon={Sparkles}
          accent
        />
      </div>
      <article className="overflow-hidden rounded-2xl border border-[#dce5e1] bg-white">
        <div className="border-b border-[#e4ebe8] px-5 py-4">
          <h2 className="font-semibold">对象排名与进度</h2>
        </div>
        <ResultTable results={results} detailed />
      </article>
    </div>
  );
}

function ResultTable({
  results,
  detailed,
}: {
  results: any[];
  detailed?: boolean;
}) {
  if (!results.length)
    return (
      <div className="p-10 text-center text-sm text-[#718078]">
        暂无评价结果
      </div>
    );
  return (
    <div className="overflow-x-auto">
      <table className="w-full text-left text-sm">
        <thead className="bg-[#f8faf9] text-xs text-[#718078]">
          <tr>
            <th className="px-5 py-3 font-medium">排名 / 对象</th>
            <th className="px-5 py-3 font-medium">评价进度</th>
            <th className="px-5 py-3 font-medium">综合得分</th>
            {detailed && <th className="px-5 py-3 font-medium">得分表现</th>}
          </tr>
        </thead>
        <tbody>
          {results.map((r, i) => (
            <tr
              key={`${r.subject}-${r.department}`}
              className="border-t border-[#edf1ef]"
            >
              <td className="px-5 py-4">
                <div className="flex items-center gap-3">
                  <span
                    className={`grid size-8 place-items-center rounded-lg text-xs font-bold ${i < 3 ? 'bg-[#d9ff70]' : 'bg-[#edf1ef]'}`}
                  >
                    {i + 1}
                  </span>
                  <span className="grid size-9 place-items-center rounded-full bg-[#eaf5f0] font-semibold text-[#1c6b56]">
                    {r.subject[0]}
                  </span>
                  <div>
                    <p className="font-medium">{r.subject}</p>
                    <p className="text-xs text-[#839089]">{r.department}</p>
                  </div>
                </div>
              </td>
              <td className="px-5 py-4 text-[#64736b]">
                {r.submitted} / {r.total}
              </td>
              <td className="px-5 py-4 text-lg font-semibold">
                {r.score ? r.score.toFixed(1) : '—'}
              </td>
              {detailed && (
                <td className="min-w-44 px-5 py-4">
                  <div className="h-2 rounded-full bg-[#edf1ef]">
                    <div
                      className="h-full rounded-full bg-[#78a58f]"
                      style={{ width: `${r.score}%` }}
                    />
                  </div>
                </td>
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function ReviewerPortal({
  data,
  onOpen,
}: {
  data: AppData;
  onOpen: (t: string, a: string) => void;
}) {
  const [loginCode, setLoginCode] = useState('');
  const [reviewerId, setReviewerId] = useState('');
  const [loginError, setLoginError] = useState('');
  useEffect(() => {
    const saved = sessionStorage.getItem('pingjian-reviewer-id');
    if (saved && data.reviewers.some((r) => r.id === saved && r.active))
      setReviewerId(saved);
  }, [data.reviewers]);
  const reviewer = data.reviewers.find((r) => r.id === reviewerId && r.active);
  const login = () => {
    const matched = data.reviewers.find(
      (r) =>
        r.active && r.code.toUpperCase() === loginCode.trim().toUpperCase(),
    );
    if (!matched) {
      setLoginError('登录码不正确或该评价人已停用');
      return;
    }
    setReviewerId(matched.id);
    sessionStorage.setItem('pingjian-reviewer-id', matched.id);
    setLoginError('');
  };
  const logout = () => {
    setReviewerId('');
    setLoginCode('');
    sessionStorage.removeItem('pingjian-reviewer-id');
  };
  if (!reviewer)
    return (
      <div className="mx-auto max-w-md pt-10">
        <article className="rounded-3xl border border-[#dce5e1] bg-white p-7 shadow-[0_18px_60px_rgba(23,48,40,.08)]">
          <span className="grid size-12 place-items-center rounded-2xl bg-[#123c32] text-[#d9ff70]">
            <KeyRound className="size-5" />
          </span>
          <h2 className="mt-5 text-xl font-semibold">评价人登录</h2>
          <p className="mt-2 text-sm leading-6 text-[#718078]">
            请输入管理员分配给您的专属登录码，登录后只显示分配给您的评价对象。
          </p>
          <label className="mt-5 block text-sm font-medium">
            专属登录码
            <Input
              value={loginCode}
              onChange={(e) => setLoginCode(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') login();
              }}
              className="mt-1 h-11 uppercase tracking-widest"
              placeholder="例如：PJ-A8K2M6"
            />
          </label>
          {loginError && (
            <p className="mt-2 text-sm text-red-600">{loginError}</p>
          )}
          <Button onClick={login} className="mt-5 h-11 w-full bg-[#123c32]">
            <LogIn />
            登录并查看任务
          </Button>
        </article>
      </div>
    );
  const ownAssignments = data.tasks
    .flatMap((t) => t.assignments.map((a) => ({ t, a })))
    .filter(
      ({ a }) =>
        a.reviewerId === reviewer.id ||
        (!a.reviewerId && a.reviewer === reviewer.name),
    );
  const pending = ownAssignments.filter(({ a }) => a.status === 'pending');
  const completed = ownAssignments.length - pending.length;
  return (
    <div className="mx-auto max-w-3xl space-y-5">
      <div className="flex items-end justify-between">
        <div>
          <p className="text-sm text-[#718078]">
            欢迎回来，{reviewer.department}
          </p>
          <h2 className="mt-1 text-xl font-semibold">
            {reviewer.name}的待评价对象
          </h2>
          <p className="mt-1 text-sm text-[#718078]">
            待完成 {pending.length} 份 · 已完成 {completed} 份
          </p>
        </div>
        <Button variant="outline" onClick={logout}>
          <LogOut />
          退出登录
        </Button>
      </div>
      {pending.length ? (
        pending.map(({ t, a }) => (
          <button
            key={a.id}
            onClick={() => onOpen(t.id, a.id)}
            className="group flex w-full items-center gap-4 rounded-2xl border border-[#dce5e1] bg-white p-5 text-left transition hover:-translate-y-0.5 hover:border-[#9db7ab] hover:shadow-lg"
          >
            <span className="grid size-11 place-items-center rounded-xl bg-[#eaf5f0] text-[#1c6b56]">
              <PencilLine className="size-5" />
            </span>
            <div className="min-w-0 flex-1">
              <h3 className="font-semibold">评价 {a.subject}</h3>
              <p className="mt-1 truncate text-sm text-[#718078]">
                {a.department} · {t.name} · 截止 {t.deadline}
              </p>
            </div>
            <span className="hidden text-xs text-[#a2781c] sm:inline">
              待填写
            </span>
            <ChevronRight className="size-5 text-[#9aa69f] transition group-hover:translate-x-1" />
          </button>
        ))
      ) : (
        <div className="rounded-2xl border border-dashed border-[#cbd8d2] bg-white p-12 text-center">
          <Check className="mx-auto size-8 text-[#1c765b]" />
          <h3 className="mt-3 font-semibold">所有评价都已完成</h3>
          <p className="mt-1 text-sm text-[#718078]">
            当前没有分配给您的待评价对象。
          </p>
        </div>
      )}
    </div>
  );
}

function PageIntro({
  title,
  text,
  action,
}: {
  title: string;
  text: string;
  action?: React.ReactNode;
}) {
  return (
    <div className="flex flex-col justify-between gap-3 md:flex-row md:items-end">
      <div>
        <h2 className="text-xl font-semibold tracking-tight">{title}</h2>
        <p className="mt-1 text-sm text-[#718078]">{text}</p>
      </div>
      {action}
    </div>
  );
}

type SheetSelection = {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
};

const cellKey = (row: number, col: number) => `${row}:${col}`;

function createDesignerTemplate() {
  const criteria: Criterion[] = [
    {
      id: `designer-${uid()}`,
      name: '工作成果',
      dimension: '原表评分项',
      weight: 60,
      max: 60,
      description: '',
    },
    {
      id: `designer-${uid()}`,
      name: '关键能力',
      dimension: '原表评分项',
      weight: 30,
      max: 30,
      description: '',
    },
    {
      id: `designer-${uid()}`,
      name: '现场答辩',
      dimension: '原表评分项',
      weight: 10,
      max: 10,
      description: '',
    },
  ];
  const merges = [
    [0, 0, 1, 6],
    [1, 0, 1, 2],
    [1, 2, 1, 2],
    [1, 4, 1, 2],
    [2, 0, 1, 2],
    [3, 0, 1, 2],
    [4, 0, 1, 2],
    [5, 0, 1, 2],
    [6, 0, 1, 2],
    [6, 2, 1, 4],
    [7, 0, 1, 2],
    [7, 2, 1, 4],
  ] as const;
  const covered = new Set<string>();
  for (const [row, col, rowSpan, colSpan] of merges) {
    for (let r = row; r < row + rowSpan; r += 1)
      for (let c = col; c < col + colSpan; c += 1)
        if (r !== row || c !== col) covered.add(cellKey(r, c));
  }
  const texts: Record<string, string> = {
    '0:0': '评价表',
    '1:0': '被评价对象：',
    '1:2': '评价人：',
    '1:4': '考察期：',
    '2:0': '考核评分要素',
    '2:2': criteria[0].name,
    '2:3': criteria[1].name,
    '2:4': criteria[2].name,
    '2:5': '备注',
    '3:0': '分值',
    '3:2': String(criteria[0].max),
    '3:3': String(criteria[1].max),
    '3:4': String(criteria[2].max),
    '4:0': '评分要点',
    '4:2': '填写评价标准和说明',
    '4:3': '填写评价标准和说明',
    '4:4': '填写评价标准和说明',
    '5:0': '合计',
    '6:0': '综合得分',
    '7:0': '综合评价',
  };
  const roles: Record<string, SheetCellRole> = {
    '1:0': 'subject',
    '1:2': 'reviewer',
    '1:4': 'stage',
    '4:5': 'remarks',
    '5:2': 'score',
    '5:3': 'score',
    '5:4': 'score',
    '6:2': 'total',
    '7:2': 'comment',
  };
  const criterionIds: Record<string, string> = {
    '5:2': criteria[0].id,
    '5:3': criteria[1].id,
    '5:4': criteria[2].id,
  };
  const mergeByStart = new Map(
    merges.map(([row, col, rowSpan, colSpan]) => [
      cellKey(row, col),
      { rowSpan, colSpan },
    ]),
  );
  const cells: SheetCell[] = [];
  for (let row = 0; row < 8; row += 1) {
    for (let col = 0; col < 6; col += 1) {
      if (covered.has(cellKey(row, col))) continue;
      const merge = mergeByStart.get(cellKey(row, col));
      cells.push({
        row,
        col,
        rowSpan: merge?.rowSpan || 1,
        colSpan: merge?.colSpan || 1,
        text: texts[cellKey(row, col)] || '',
        role: roles[cellKey(row, col)],
        criterionId: criterionIds[cellKey(row, col)],
        style: {
          bold: row === 0 || row === 2 || row === 3 || col < 2,
          fontSize: row === 0 ? 26 : 14,
          align: row === 4 && col >= 2 && col <= 4 ? 'left' : 'center',
          vertical: 'middle',
          wrap: true,
          background: row === 2 || row === 3 ? '#f3f6f4' : undefined,
        },
      });
    }
  }
  return {
    criteria,
    layout: {
      sheetName: '在线表格',
      rowCount: 8,
      colCount: 6,
      columnWidths: [90, 90, 230, 230, 210, 150],
      rowHeights: [68, 52, 52, 48, 190, 62, 66, 110],
      cells,
    } satisfies SheetLayout,
  };
}

const roleLabels: Array<{ value: SheetCellRole | ''; label: string }> = [
  { value: '', label: '普通文字' },
  { value: 'subject', label: '被评价对象' },
  { value: 'reviewer', label: '评价人' },
  { value: 'stage', label: '考察期' },
  { value: 'score', label: '评分输入框' },
  { value: 'input', label: '普通填写框' },
  { value: 'remarks', label: '备注输入框' },
  { value: 'comment', label: '综合评价输入框' },
  { value: 'total', label: '自动综合得分' },
  { value: 'signature', label: '签名和日期' },
];

function SheetDesigner({
  layout,
  setLayout,
  criteria,
  setCriteria,
  setNotice,
}: {
  layout: SheetLayout;
  setLayout: React.Dispatch<React.SetStateAction<SheetLayout | undefined>>;
  criteria: Criterion[];
  setCriteria: React.Dispatch<React.SetStateAction<Criterion[]>>;
  setNotice: (message: string) => void;
}) {
  const [selection, setSelection] = useState<SheetSelection>({
    startRow: 0,
    startCol: 0,
    endRow: 0,
    endCol: 0,
  });
  const bounds = {
    top: Math.min(selection.startRow, selection.endRow),
    bottom: Math.max(selection.startRow, selection.endRow),
    left: Math.min(selection.startCol, selection.endCol),
    right: Math.max(selection.startCol, selection.endCol),
  };
  const cellMap = new Map(
    layout.cells.map((cell) => [cellKey(cell.row, cell.col), cell]),
  );
  const selectedCell = cellMap.get(cellKey(bounds.top, bounds.left));
  const selectedCriterion = criteria.find(
    (criterion) => criterion.id === selectedCell?.criterionId,
  );
  const updateCell = (patch: Partial<SheetCell>) =>
    setLayout((current) =>
      current
        ? {
            ...current,
            cells: current.cells.map((cell) =>
              cell.row === bounds.top && cell.col === bounds.left
                ? { ...cell, ...patch }
                : cell,
            ),
          }
        : current,
    );
  const setRole = (role: SheetCellRole | '') => {
    if (!selectedCell) return;
    if (role === 'score') {
      const criterion =
        selectedCriterion ||
        ({
          id: `designer-${uid()}`,
          name: selectedCell.text || '评分项',
          dimension: '自定义表格',
          weight: 0,
          max: 100,
          description: '',
        } satisfies Criterion);
      if (!selectedCriterion) setCriteria((current) => [...current, criterion]);
      updateCell({ role, criterionId: criterion.id, fieldId: undefined });
    } else
      updateCell({
        role: role || undefined,
        criterionId: undefined,
        fieldId:
          role === 'input'
            ? selectedCell.fieldId || `field-${uid()}`
            : undefined,
      });
  };
  const addRow = () =>
    setLayout((current) => {
      if (!current) return current;
      const row = current.rowCount;
      return {
        ...current,
        rowCount: row + 1,
        rowHeights: [...current.rowHeights, 48],
        cells: [
          ...current.cells,
          ...Array.from({ length: current.colCount }, (_, col) => ({
            row,
            col,
            rowSpan: 1,
            colSpan: 1,
            text: '',
            style: { align: 'center' as const, vertical: 'middle' as const },
          })),
        ],
      };
    });
  const addColumn = () =>
    setLayout((current) => {
      if (!current) return current;
      const col = current.colCount;
      return {
        ...current,
        colCount: col + 1,
        columnWidths: [...current.columnWidths, 140],
        cells: [
          ...current.cells,
          ...Array.from({ length: current.rowCount }, (_, row) => ({
            row,
            col,
            rowSpan: 1,
            colSpan: 1,
            text: '',
            style: { align: 'center' as const, vertical: 'middle' as const },
          })),
        ],
      };
    });
  const deleteRow = () => {
    if (layout.rowCount <= 1) return;
    const rowToDelete = bounds.top;
    if (
      layout.cells.some(
        (cell) =>
          cell.rowSpan > 1 &&
          cell.row <= rowToDelete &&
          cell.row + cell.rowSpan - 1 >= rowToDelete,
      )
    ) {
      setNotice('该行包含合并单元格，请先取消合并');
      return;
    }
    setLayout((current) =>
      current
        ? {
            ...current,
            rowCount: current.rowCount - 1,
            rowHeights: current.rowHeights.filter(
              (_, index) => index !== rowToDelete,
            ),
            cells: current.cells
              .filter((cell) => cell.row !== rowToDelete)
              .map((cell) =>
                cell.row > rowToDelete ? { ...cell, row: cell.row - 1 } : cell,
              ),
          }
        : current,
    );
    const nextRow = Math.max(0, Math.min(rowToDelete, layout.rowCount - 2));
    setSelection({
      startRow: nextRow,
      endRow: nextRow,
      startCol: bounds.left,
      endCol: bounds.left,
    });
  };
  const deleteColumn = () => {
    if (layout.colCount <= 1) return;
    const colToDelete = bounds.left;
    if (
      layout.cells.some(
        (cell) =>
          cell.colSpan > 1 &&
          cell.col <= colToDelete &&
          cell.col + cell.colSpan - 1 >= colToDelete,
      )
    ) {
      setNotice('该列包含合并单元格，请先取消合并');
      return;
    }
    setLayout((current) =>
      current
        ? {
            ...current,
            colCount: current.colCount - 1,
            columnWidths: current.columnWidths.filter(
              (_, index) => index !== colToDelete,
            ),
            cells: current.cells
              .filter((cell) => cell.col !== colToDelete)
              .map((cell) =>
                cell.col > colToDelete ? { ...cell, col: cell.col - 1 } : cell,
              ),
          }
        : current,
    );
    const nextCol = Math.max(0, Math.min(colToDelete, layout.colCount - 2));
    setSelection({
      startRow: bounds.top,
      endRow: bounds.top,
      startCol: nextCol,
      endCol: nextCol,
    });
  };
  const mergeSelection = () => {
    if (bounds.top === bounds.bottom && bounds.left === bounds.right) {
      setNotice('请按住 Shift 再点击另一个单元格，选择需要合并的区域');
      return;
    }
    const intersectsMerged = layout.cells.some(
      (cell) =>
        (cell.rowSpan > 1 || cell.colSpan > 1) &&
        cell.row <= bounds.bottom &&
        cell.row + cell.rowSpan - 1 >= bounds.top &&
        cell.col <= bounds.right &&
        cell.col + cell.colSpan - 1 >= bounds.left,
    );
    if (intersectsMerged) {
      setNotice('所选区域包含已合并单元格，请先选中该单元格并取消合并');
      return;
    }
    setLayout((current) => {
      if (!current) return current;
      const selected = current.cells.filter(
        (cell) =>
          cell.row >= bounds.top &&
          cell.row <= bounds.bottom &&
          cell.col >= bounds.left &&
          cell.col <= bounds.right,
      );
      const first = selected.find(
        (cell) => cell.row === bounds.top && cell.col === bounds.left,
      );
      return {
        ...current,
        cells: [
          ...current.cells.filter(
            (cell) =>
              !(
                cell.row >= bounds.top &&
                cell.row <= bounds.bottom &&
                cell.col >= bounds.left &&
                cell.col <= bounds.right
              ),
          ),
          {
            ...(first || {
              row: bounds.top,
              col: bounds.left,
              text: '',
            }),
            row: bounds.top,
            col: bounds.left,
            rowSpan: bounds.bottom - bounds.top + 1,
            colSpan: bounds.right - bounds.left + 1,
            text: selected
              .map((cell) => cell.text)
              .filter(Boolean)
              .join(' '),
          },
        ],
      };
    });
    setSelection({
      startRow: bounds.top,
      startCol: bounds.left,
      endRow: bounds.top,
      endCol: bounds.left,
    });
  };
  const unmerge = () => {
    if (
      !selectedCell ||
      (selectedCell.rowSpan === 1 && selectedCell.colSpan === 1)
    ) {
      setNotice('当前单元格没有合并');
      return;
    }
    setLayout((current) => {
      if (!current) return current;
      const additions: SheetCell[] = [];
      for (
        let row = selectedCell.row;
        row < selectedCell.row + selectedCell.rowSpan;
        row += 1
      )
        for (
          let col = selectedCell.col;
          col < selectedCell.col + selectedCell.colSpan;
          col += 1
        )
          additions.push({
            row,
            col,
            rowSpan: 1,
            colSpan: 1,
            text:
              row === selectedCell.row && col === selectedCell.col
                ? selectedCell.text
                : '',
            role:
              row === selectedCell.row && col === selectedCell.col
                ? selectedCell.role
                : undefined,
            criterionId:
              row === selectedCell.row && col === selectedCell.col
                ? selectedCell.criterionId
                : undefined,
            style: selectedCell.style,
          });
      return {
        ...current,
        cells: [
          ...current.cells.filter(
            (cell) =>
              cell.row !== selectedCell.row || cell.col !== selectedCell.col,
          ),
          ...additions,
        ],
      };
    });
  };
  const updateCriterion = (patch: Partial<Criterion>) => {
    if (!selectedCriterion) return;
    setCriteria((current) =>
      current.map((criterion) =>
        criterion.id === selectedCriterion.id
          ? { ...criterion, ...patch }
          : criterion,
      ),
    );
  };
  const resizeColumn = (col: number, event: React.MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    const startX = event.clientX;
    const startWidth = layout.columnWidths[col];
    const move = (moveEvent: MouseEvent) => {
      const nextWidth = Math.max(50, startWidth + moveEvent.clientX - startX);
      setLayout((current) => {
        if (!current) return current;
        const widths = [...current.columnWidths];
        widths[col] = nextWidth;
        return { ...current, columnWidths: widths };
      });
    };
    const stop = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', stop);
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', stop);
  };
  const resizeRow = (row: number, event: React.MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    const startY = event.clientY;
    const startHeight = layout.rowHeights[row];
    const move = (moveEvent: MouseEvent) => {
      const nextHeight = Math.max(32, startHeight + moveEvent.clientY - startY);
      setLayout((current) => {
        if (!current) return current;
        const heights = [...current.rowHeights];
        heights[row] = nextHeight;
        return { ...current, rowHeights: heights };
      });
    };
    const stop = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', stop);
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', stop);
  };
  const columnLabel = (col: number) => {
    let label = '';
    let value = col + 1;
    while (value > 0) {
      value -= 1;
      label = String.fromCharCode(65 + (value % 26)) + label;
      value = Math.floor(value / 26);
    }
    return label;
  };
  return (
    <div className="space-y-3">
      <div className="flex flex-wrap items-center gap-2 rounded-xl border border-[#dce5e1] bg-[#f7faf8] p-2">
        <Button type="button" size="sm" variant="outline" onClick={addRow}>
          <Plus /> 添加行
        </Button>
        <Button type="button" size="sm" variant="outline" onClick={addColumn}>
          <Plus /> 添加列
        </Button>
        <Button type="button" size="sm" variant="outline" onClick={deleteRow}>
          删除选中行
        </Button>
        <Button
          type="button"
          size="sm"
          variant="outline"
          onClick={deleteColumn}
        >
          删除选中列
        </Button>
        <Button
          type="button"
          size="sm"
          variant="outline"
          onClick={mergeSelection}
        >
          合并所选区域
        </Button>
        <Button type="button" size="sm" variant="outline" onClick={unmerge}>
          取消合并
        </Button>
        <span className="ml-auto text-xs text-[#66736d]">
          点击选择；按住 Shift 点击可框选区域
        </span>
      </div>
      <div className="grid gap-3 xl:grid-cols-[minmax(0,1fr)_260px]">
        <div className="overflow-auto rounded-lg border border-[#7a8580] bg-white">
          <table
            className="table-fixed border-collapse text-sm"
            style={{
              width:
                42 + layout.columnWidths.reduce((sum, width) => sum + width, 0),
            }}
          >
            <colgroup>
              <col style={{ width: 42 }} />
              {layout.columnWidths.map((width, col) => (
                <col key={col} style={{ width }} />
              ))}
            </colgroup>
            <thead>
              <tr className="h-8 bg-[#e9eeeb] text-xs text-[#52615a]">
                <th className="sticky left-0 z-20 border border-[#8d9892] bg-[#dfe6e2]" />
                {layout.columnWidths.map((_, col) => (
                  <th
                    key={col}
                    className="relative select-none border border-[#8d9892] text-center font-medium"
                    onClick={() =>
                      setSelection({
                        startRow: 0,
                        endRow: 0,
                        startCol: col,
                        endCol: col,
                      })
                    }
                  >
                    {columnLabel(col)}
                    <button
                      type="button"
                      aria-label={`调整第 ${col + 1} 列宽度`}
                      className="absolute -right-1 top-0 z-10 h-full w-2 cursor-col-resize hover:bg-[#1c8065]/30"
                      onMouseDown={(event) => resizeColumn(col, event)}
                    />
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {Array.from({ length: layout.rowCount }, (_, row) => (
                <tr key={row} style={{ height: layout.rowHeights[row] }}>
                  <th
                    className="sticky left-0 z-10 select-none border border-[#8d9892] bg-[#e9eeeb] text-center text-xs font-medium text-[#52615a]"
                    onClick={() =>
                      setSelection({
                        startRow: row,
                        endRow: row,
                        startCol: 0,
                        endCol: 0,
                      })
                    }
                  >
                    {row + 1}
                    <button
                      type="button"
                      aria-label={`调整第 ${row + 1} 行高度`}
                      className="absolute -bottom-1 left-0 z-10 h-2 w-full cursor-row-resize hover:bg-[#1c8065]/30"
                      onMouseDown={(event) => resizeRow(row, event)}
                    />
                  </th>
                  {Array.from({ length: layout.colCount }, (_, col) => {
                    const cell = cellMap.get(cellKey(row, col));
                    if (!cell) return null;
                    const selected =
                      cell.row <= bounds.bottom &&
                      cell.row + cell.rowSpan - 1 >= bounds.top &&
                      cell.col <= bounds.right &&
                      cell.col + cell.colSpan - 1 >= bounds.left;
                    return (
                      <td
                        key={col}
                        rowSpan={cell.rowSpan}
                        colSpan={cell.colSpan}
                        className={`relative border border-[#4c5651] p-0 ${
                          selected
                            ? 'outline outline-2 -outline-offset-2 outline-[#1d8065]'
                            : ''
                        }`}
                        style={{ backgroundColor: cell.style?.background }}
                      >
                        <textarea
                          value={cell.text}
                          onMouseDown={(event) => {
                            if (event.shiftKey) {
                              event.preventDefault();
                              setSelection((current) => ({
                                ...current,
                                endRow: cell.row,
                                endCol: cell.col,
                              }));
                            } else {
                              setSelection({
                                startRow: cell.row,
                                startCol: cell.col,
                                endRow: cell.row,
                                endCol: cell.col,
                              });
                            }
                          }}
                          onChange={(event) => {
                            const value = event.target.value;
                            setLayout((current) =>
                              current
                                ? {
                                    ...current,
                                    cells: current.cells.map((item) =>
                                      item.row === cell.row &&
                                      item.col === cell.col
                                        ? { ...item, text: value }
                                        : item,
                                    ),
                                  }
                                : current,
                            );
                          }}
                          className="size-full min-h-10 resize-none bg-transparent px-2 py-1.5 outline-none"
                          style={{
                            fontWeight: cell.style?.bold ? 700 : 400,
                            fontSize: cell.style?.fontSize,
                            textAlign: cell.style?.align || 'left',
                          }}
                          aria-label={`第 ${row + 1} 行第 ${col + 1} 列`}
                        />
                        {cell.role && (
                          <span className="pointer-events-none absolute right-1 top-1 rounded bg-[#123c32] px-1.5 py-0.5 text-[9px] text-white">
                            {
                              roleLabels.find(
                                (item) => item.value === cell.role,
                              )?.label
                            }
                          </span>
                        )}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <aside className="space-y-3 rounded-xl border border-[#dce5e1] bg-white p-3">
          <div>
            <p className="text-sm font-semibold">单元格设置</p>
            <p className="text-xs text-[#718078]">
              第 {bounds.top + 1} 行，第 {bounds.left + 1} 列
            </p>
          </div>
          <label className="block text-xs font-medium">
            字段类型
            <select
              value={selectedCell?.role || ''}
              onChange={(event) =>
                setRole(event.target.value as SheetCellRole | '')
              }
              className="mt-1 h-9 w-full rounded-lg border border-[#cfd9d4] bg-white px-2"
            >
              {roleLabels.map((item) => (
                <option key={item.value || 'text'} value={item.value}>
                  {item.label}
                </option>
              ))}
            </select>
          </label>
          <div className="grid grid-cols-3 gap-1">
            {(['left', 'center', 'right'] as const).map((align) => (
              <Button
                type="button"
                key={align}
                size="sm"
                variant={
                  selectedCell?.style?.align === align ? 'default' : 'outline'
                }
                onClick={() =>
                  updateCell({
                    style: { ...selectedCell?.style, align },
                  })
                }
              >
                {align === 'left' ? '左' : align === 'center' ? '中' : '右'}
              </Button>
            ))}
          </div>
          <Button
            type="button"
            size="sm"
            variant={selectedCell?.style?.bold ? 'default' : 'outline'}
            className="w-full"
            onClick={() =>
              updateCell({
                style: {
                  ...selectedCell?.style,
                  bold: !selectedCell?.style?.bold,
                },
              })
            }
          >
            加粗
          </Button>
          <div className="grid grid-cols-2 gap-2">
            <label className="text-xs font-medium">
              字号
              <Input
                type="number"
                min={10}
                max={48}
                value={selectedCell?.style?.fontSize || 14}
                onChange={(event) =>
                  updateCell({
                    style: {
                      ...selectedCell?.style,
                      fontSize: Number(event.target.value),
                    },
                  })
                }
              />
            </label>
            <label className="text-xs font-medium">
              背景色
              <Input
                type="color"
                value={selectedCell?.style?.background || '#ffffff'}
                onChange={(event) =>
                  updateCell({
                    style: {
                      ...selectedCell?.style,
                      background: event.target.value,
                    },
                  })
                }
                className="p-1"
              />
            </label>
          </div>
          <div className="grid grid-cols-2 gap-2">
            <label className="text-xs font-medium">
              列宽（列距）
              <Input
                type="number"
                min={50}
                value={layout.columnWidths[bounds.left]}
                onChange={(event) =>
                  setLayout((current) => {
                    if (!current) return current;
                    const widths = [...current.columnWidths];
                    widths[bounds.left] = Math.max(
                      50,
                      Number(event.target.value),
                    );
                    return { ...current, columnWidths: widths };
                  })
                }
              />
            </label>
            <label className="text-xs font-medium">
              行高（行距）
              <Input
                type="number"
                min={32}
                value={layout.rowHeights[bounds.top]}
                onChange={(event) =>
                  setLayout((current) => {
                    if (!current) return current;
                    const heights = [...current.rowHeights];
                    heights[bounds.top] = Math.max(
                      32,
                      Number(event.target.value),
                    );
                    return { ...current, rowHeights: heights };
                  })
                }
              />
            </label>
          </div>
          {selectedCell?.role === 'score' && selectedCriterion && (
            <div className="space-y-2 border-t border-[#e5ebe8] pt-3">
              <p className="text-xs font-semibold text-[#185b49]">评分字段</p>
              <label className="block text-xs font-medium">
                字段名称
                <Input
                  value={selectedCriterion.name}
                  onChange={(event) =>
                    updateCriterion({ name: event.target.value })
                  }
                />
              </label>
              <div className="grid grid-cols-2 gap-2">
                <label className="text-xs font-medium">
                  满分
                  <Input
                    type="number"
                    min={1}
                    value={selectedCriterion.max}
                    onChange={(event) =>
                      updateCriterion({ max: Number(event.target.value) })
                    }
                  />
                </label>
                <label className="text-xs font-medium">
                  权重
                  <Input
                    type="number"
                    min={0}
                    value={selectedCriterion.weight}
                    onChange={(event) =>
                      updateCriterion({ weight: Number(event.target.value) })
                    }
                  />
                </label>
              </div>
            </div>
          )}
        </aside>
      </div>
    </div>
  );
}

function TemplateDialog({
  open,
  onOpenChange,
  setData,
  setNotice,
  template,
}: {
  open: boolean;
  onOpenChange: (x: boolean) => void;
  setData: React.Dispatch<React.SetStateAction<AppData>>;
  setNotice: (s: string) => void;
  template: Template | null;
}) {
  const fileRef = useRef<HTMLInputElement>(null);
  const [initialDesigner] = useState(createDesignerTemplate);
  const [name, setName] = useState('');
  const [description, setDescription] = useState('');
  const [editorMode, setEditorMode] = useState<'sheet' | 'criteria'>('sheet');
  const [sheetLayout, setSheetLayout] = useState<SheetLayout | undefined>(
    initialDesigner.layout,
  );
  const [criteria, setCriteria] = useState<Criterion[]>(
    initialDesigner.criteria,
  );
  useEffect(() => {
    if (!open) return;
    if (template) {
      setName(template.name);
      setDescription(template.description);
      setCriteria(template.criteria.map((criterion) => ({ ...criterion })));
      setSheetLayout(
        template.sheetLayout
          ? {
              ...template.sheetLayout,
              columnWidths: [...template.sheetLayout.columnWidths],
              rowHeights: [...template.sheetLayout.rowHeights],
              cells: template.sheetLayout.cells.map((cell) => ({
                ...cell,
                style: cell.style ? { ...cell.style } : undefined,
              })),
            }
          : undefined,
      );
      setEditorMode(template.sheetLayout ? 'sheet' : 'criteria');
      return;
    }
    const fresh = createDesignerTemplate();
    setName('');
    setDescription('');
    setCriteria(fresh.criteria);
    setSheetLayout(fresh.layout);
    setEditorMode('sheet');
  }, [open, template]);
  const add = () =>
    setCriteria((c) => [
      ...c,
      {
        id: uid(),
        name: '',
        dimension: '综合表现',
        weight: 0,
        max: 100,
        description: '',
      },
    ]);
  const update = (id: string, key: keyof Criterion, value: string | number) =>
    setCriteria((c) =>
      c.map((x) => (x.id === id ? { ...x, [key]: value } : x)),
    );
  const importFile = async (file?: File) => {
    if (!file) return;
    const wb = XLSX.read(await file.arrayBuffer(), {
      cellStyles: true,
      cellFormula: false,
    });
    if (/\.xlsx?$/i.test(file.name)) {
      const parsed = parseSheetLayout(wb);
      if (parsed) {
        setCriteria(parsed.criteria);
        setSheetLayout(parsed.layout);
        setEditorMode('sheet');
        if (!name) setName(file.name.replace(/\.xlsx?$/i, ''));
        setDescription((current) => current || '按原 Excel 表格结构生成');
        setNotice(
          `已保留原表 ${parsed.layout.rowCount} 行 × ${parsed.layout.colCount} 列结构，并识别 ${parsed.criteria.length} 个评分项`,
        );
        return;
      }
    }
    const rows = XLSX.utils.sheet_to_json<Record<string, any>>(
      wb.Sheets[wb.SheetNames[0]],
      { defval: '' },
    );
    if (!rows.length) {
      setNotice('未识别到有效表格内容');
      return;
    }
    const firstKeys = Object.keys(rows[0]);
    const find = (keys: string[]) =>
      firstKeys.find((k) => keys.some((x) => k.toLowerCase().includes(x))) ||
      '';
    const nameKey =
      find(['指标', '项目', '内容', 'criterion', 'item']) || firstKeys[0];
    const dimKey = find(['维度', '分类', 'dimension']);
    const weightKey = find(['权重', 'weight']);
    const maxKey = find(['满分', '分值', 'max']);
    const descKey = find(['说明', '标准', 'description']);
    const imported = rows
      .filter((r) => String(r[nameKey]).trim())
      .map((r) => ({
        id: uid(),
        name: String(r[nameKey]).trim(),
        dimension: dimKey ? String(r[dimKey] || '综合表现') : '综合表现',
        weight: weightKey
          ? Number(r[weightKey]) || 0
          : Math.round(100 / rows.length),
        max: maxKey ? Number(r[maxKey]) || 100 : 100,
        description: descKey ? String(r[descKey] || '') : '',
      }));
    setCriteria(imported);
    setSheetLayout(undefined);
    setEditorMode('criteria');
    if (!name) setName(file.name.replace(/\.(xlsx?|csv)$/i, ''));
    setNotice(`已识别 ${imported.length} 个评价指标`);
  };
  const save = () => {
    const scoreIds = new Set(
      sheetLayout?.cells
        .filter((cell) => cell.role === 'score' && cell.criterionId)
        .map((cell) => cell.criterionId!) || [],
    );
    const valid = criteria.filter(
      (criterion) =>
        criterion.name.trim() &&
        (editorMode === 'criteria' || scoreIds.has(criterion.id)),
    );
    if (!name.trim() || !valid.length) {
      setNotice(
        editorMode === 'sheet'
          ? '请填写模板名称，并至少设置一个“评分输入框”单元格'
          : '请填写模板名称和至少一个指标',
      );
      return;
    }
    setData((d) => ({
      ...d,
      templates: [
        ...d.templates,
        {
          id: uid(),
          familyId: template?.familyId || template?.id,
          name: name.trim(),
          description: description.trim() || '自定义评价表',
          criteria: valid,
          updatedAt: today(),
          sheetLayout: editorMode === 'sheet' ? sheetLayout : undefined,
          version: template ? (template.version || 1) + 1 : 1,
        },
      ],
    }));
    onOpenChange(false);
    setName('');
    setDescription('');
    const freshDesigner = createDesignerTemplate();
    setEditorMode('sheet');
    setSheetLayout(freshDesigner.layout);
    setCriteria(freshDesigner.criteria);
    setNotice(
      template
        ? `已保存为新版本 v${(template.version || 1) + 1}，原任务仍使用旧版本`
        : '评价表模板已保存',
    );
  };
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-h-[94vh] overflow-y-auto sm:max-w-[min(96vw,1540px)]">
        <DialogHeader>
          <DialogTitle>
            {template
              ? `复制编辑模板 · 新版本 v${(template.version || 1) + 1}`
              : '新建评价表模板'}
          </DialogTitle>
          <DialogDescription>
            在平台内像编辑 Excel 一样设计评价表，也可以上传 Excel
            作为可继续修改的草稿。
          </DialogDescription>
        </DialogHeader>
        <div className="space-y-4">
          <div className="grid gap-3 sm:grid-cols-2">
            <label className="text-sm font-medium">
              模板名称
              <Input
                className="mt-1 h-10"
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="例如：年度管理能力评价表"
              />
            </label>
            <label className="text-sm font-medium">
              模板说明
              <Input
                className="mt-1 h-10"
                value={description}
                onChange={(e) => setDescription(e.target.value)}
                placeholder="适用范围或评价说明"
              />
            </label>
          </div>
          <div className="flex flex-wrap items-center gap-2 rounded-xl bg-[#edf2ef] p-1.5">
            <Button
              type="button"
              size="sm"
              variant={editorMode === 'sheet' ? 'default' : 'ghost'}
              className={editorMode === 'sheet' ? 'bg-[#123c32]' : ''}
              onClick={() => {
                if (!sheetLayout) {
                  const fresh = createDesignerTemplate();
                  setSheetLayout(fresh.layout);
                  setCriteria(fresh.criteria);
                }
                setEditorMode('sheet');
              }}
            >
              类 Excel 表格设计
            </Button>
            <Button
              type="button"
              size="sm"
              variant={editorMode === 'criteria' ? 'default' : 'ghost'}
              className={editorMode === 'criteria' ? 'bg-[#123c32]' : ''}
              onClick={() => setEditorMode('criteria')}
            >
              简单指标列表
            </Button>
            <span className="ml-auto px-2 text-xs text-[#66736d]">
              当前模板版本 v1
            </span>
          </div>
          <input
            ref={fileRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            className="hidden"
            onChange={(e) => importFile(e.target.files?.[0])}
          />
          <button
            onClick={() => fileRef.current?.click()}
            className="flex w-full items-center justify-center gap-3 rounded-xl border border-dashed border-[#9db7ab] bg-[#f4f8f6] px-4 py-4 text-sm font-medium text-[#1c6b56] hover:bg-[#eaf5f0]"
          >
            <Upload className="size-5" />
            上传 Excel 还原原表 / 上传 CSV 识别指标
            <span className="text-xs font-normal text-[#718078]">
              Excel 将保留合并单元格、行高、列宽与评分位置
            </span>
          </button>
          {editorMode === 'sheet' && sheetLayout && (
            <div className="rounded-xl border border-[#b9d8ca] bg-[#edf8f3] px-4 py-3 text-sm text-[#185b49]">
              <b>表格结构可编辑</b>
              <span className="ml-2">
                {sheetLayout.sheetName} · {sheetLayout.rowCount} 行 ×{' '}
                {sheetLayout.colCount} 列 · 上传 Excel 后仍可继续调整
              </span>
            </div>
          )}
          {editorMode === 'sheet' && sheetLayout ? (
            <SheetDesigner
              layout={sheetLayout}
              setLayout={setSheetLayout}
              criteria={criteria}
              setCriteria={setCriteria}
              setNotice={setNotice}
            />
          ) : (
            <>
              <div className="flex items-center justify-between">
                <h4 className="text-sm font-semibold">
                  评价指标{' '}
                  <span className="font-normal text-[#718078]">
                    · 权重合计{' '}
                    {criteria.reduce((s, c) => s + Number(c.weight), 0)}%
                  </span>
                </h4>
                <Button size="sm" variant="outline" onClick={add}>
                  <Plus />
                  添加指标
                </Button>
              </div>
              <div className="space-y-3">
                {criteria.map((c, i) => (
                  <div
                    key={c.id}
                    className="grid gap-2 rounded-xl border border-[#e1e9e5] bg-[#fafcfb] p-3 sm:grid-cols-[1.1fr_1.4fr_84px_84px_auto]"
                  >
                    <Input
                      value={c.dimension}
                      onChange={(e) =>
                        update(c.id, 'dimension', e.target.value)
                      }
                      placeholder="维度"
                    />
                    <Input
                      value={c.name}
                      onChange={(e) => update(c.id, 'name', e.target.value)}
                      placeholder={`指标 ${i + 1}`}
                    />
                    <Input
                      type="number"
                      value={c.weight}
                      onChange={(e) =>
                        update(c.id, 'weight', Number(e.target.value))
                      }
                      aria-label="权重"
                    />
                    <Input
                      type="number"
                      value={c.max}
                      onChange={(e) =>
                        update(c.id, 'max', Number(e.target.value))
                      }
                      aria-label="满分"
                    />
                    <button
                      onClick={() =>
                        setCriteria((x) => x.filter((y) => y.id !== c.id))
                      }
                      className="grid size-8 place-items-center text-[#9aa69f] hover:text-red-600"
                    >
                      <Trash2 className="size-4" />
                    </button>
                  </div>
                ))}
              </div>
            </>
          )}
        </div>
        <DialogFooter>
          <Button variant="outline" onClick={() => onOpenChange(false)}>
            取消
          </Button>
          <Button onClick={save} className="bg-[#123c32]">
            保存模板
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

function TaskDialog({
  open,
  onOpenChange,
  data,
  setData,
  setNotice,
}: {
  open: boolean;
  onOpenChange: (x: boolean) => void;
  data: AppData;
  setData: React.Dispatch<React.SetStateAction<AppData>>;
  setNotice: (s: string) => void;
}) {
  const [name, setName] = useState('');
  const [templateId, setTemplateId] = useState(data.templates[0]?.id || '');
  const [deadline, setDeadline] = useState('2026-09-30');
  const [stageStartDate, setStageStartDate] = useState('');
  const [stageEndDate, setStageEndDate] = useState('');
  const [subjects, setSubjects] = useState('');
  const [assignmentMap, setAssignmentMap] = useState<Record<string, string[]>>(
    {},
  );
  useEffect(() => {
    if (!templateId && data.templates[0]) setTemplateId(data.templates[0].id);
  }, [data.templates, templateId]);
  const people = useMemo(
    () =>
      subjects
        .split(/\n|,|，/)
        .map((value) => value.trim())
        .filter(Boolean)
        .map((raw) => {
          const [subject, department = '未分组'] = raw
            .split(/[|｜]/)
            .map((value) => value.trim());
          return { key: `${subject}|${department}`, subject, department };
        }),
    [subjects],
  );
  const activeReviewers = data.reviewers.filter((reviewer) => reviewer.active);
  const toggleAssignment = (
    reviewerId: string,
    subjectKey: string,
    checked: boolean,
  ) =>
    setAssignmentMap((current) => ({
      ...current,
      [reviewerId]: checked
        ? [...new Set([...(current[reviewerId] || []), subjectKey])]
        : (current[reviewerId] || []).filter((key) => key !== subjectKey),
    }));
  const save = () => {
    if (!name.trim() || !templateId || !people.length) {
      setNotice('请完整填写任务名称和被评价对象');
      return;
    }
    if (
      templateId === 'tpl-stage-report' &&
      (!stageStartDate || !stageEndDate)
    ) {
      setNotice('请选择考察阶段的开始日期和结束日期');
      return;
    }
    if (templateId === 'tpl-stage-report' && stageEndDate < stageStartDate) {
      setNotice('考察阶段结束日期不能早于开始日期');
      return;
    }
    const assignments: Assignment[] = activeReviewers.flatMap((reviewer) =>
      people
        .filter((person) => assignmentMap[reviewer.id]?.includes(person.key))
        .map((person) => ({
          id: uid(),
          subject: person.subject,
          department: person.department,
          reviewer: reviewer.name,
          reviewerId: reviewer.id,
          status: 'pending' as const,
          scores: {},
          comment: '',
        })),
    );
    if (!assignments.length) {
      setNotice('请在分配矩阵中至少勾选一项');
      return;
    }
    setData((current) => ({
      ...current,
      tasks: [
        ...current.tasks,
        {
          id: uid(),
          name: name.trim(),
          templateId,
          deadline,
          stageStartDate:
            templateId === 'tpl-stage-report' ? stageStartDate : undefined,
          stageEndDate:
            templateId === 'tpl-stage-report' ? stageEndDate : undefined,
          createdAt: today(),
          assignments,
        },
      ],
    }));
    onOpenChange(false);
    setName('');
    setStageStartDate('');
    setStageEndDate('');
    setSubjects('');
    setAssignmentMap({});
    setNotice(`任务已创建，共生成 ${assignments.length} 份评价`);
  };
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-h-[92vh] overflow-y-auto sm:max-w-4xl">
        <DialogHeader>
          <DialogTitle>新建评价任务</DialogTitle>
          <DialogDescription>
            录入对象后，在分配矩阵中为每位评价人勾选不同的评价对象。
          </DialogDescription>
        </DialogHeader>
        <div className="space-y-4">
          <label className="block text-sm font-medium">
            任务名称
            <Input
              className="mt-1 h-10"
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="例如：2026 年第四季度综合评价"
            />
          </label>
          <div
            className={`grid gap-3 ${templateId === 'tpl-stage-report' ? 'sm:grid-cols-4' : 'sm:grid-cols-2'}`}
          >
            <label className="text-sm font-medium">
              使用模板
              <select
                className="mt-1 h-10 w-full rounded-lg border border-[#d9e2de] bg-white px-3 text-sm"
                value={templateId}
                onChange={(e) => setTemplateId(e.target.value)}
              >
                {data.templates.map((t) => (
                  <option key={t.id} value={t.id}>
                    {t.name} · v{t.version || 1}
                  </option>
                ))}
              </select>
            </label>
            {templateId === 'tpl-stage-report' && (
              <label className="text-sm font-medium">
                考察阶段开始日期
                <Input
                  className="mt-1 h-10"
                  type="date"
                  value={stageStartDate}
                  onChange={(e) => setStageStartDate(e.target.value)}
                />
              </label>
            )}
            {templateId === 'tpl-stage-report' && (
              <label className="text-sm font-medium">
                考察阶段结束日期
                <Input
                  className="mt-1 h-10"
                  type="date"
                  min={stageStartDate || undefined}
                  value={stageEndDate}
                  onChange={(e) => setStageEndDate(e.target.value)}
                />
              </label>
            )}
            <label className="text-sm font-medium">
              截止日期
              <Input
                className="mt-1 h-10"
                type="date"
                value={deadline}
                onChange={(e) => setDeadline(e.target.value)}
              />
            </label>
          </div>
          <label className="block text-sm font-medium">
            被评价对象{' '}
            <span className="font-normal text-[#718078]">
              （每行一人，格式：姓名 | 部门）
            </span>
            <Textarea
              className="mt-1 min-h-24"
              value={subjects}
              onChange={(e) => setSubjects(e.target.value)}
              placeholder={'张三 | 产品部\n李四 | 运营部'}
            />
          </label>
          <div>
            <div className="mb-2 flex items-center justify-between">
              <h4 className="text-sm font-semibold">评价对象分配矩阵</h4>
              <span className="text-xs text-[#718078]">
                每个勾选项生成一份评价
              </span>
            </div>
            {!activeReviewers.length ? (
              <div className="rounded-xl border border-dashed p-6 text-center text-sm text-[#718078]">
                请先在“评价人管理”中添加并启用评价人
              </div>
            ) : !people.length ? (
              <div className="rounded-xl border border-dashed p-6 text-center text-sm text-[#718078]">
                录入被评价对象后即可分配
              </div>
            ) : (
              <div className="overflow-x-auto rounded-xl border border-[#dce5e1]">
                <table className="w-full min-w-[620px] text-sm">
                  <thead className="bg-[#f6f8f7]">
                    <tr>
                      <th className="px-3 py-3 text-left font-medium">
                        评价人
                      </th>
                      {people.map((person) => (
                        <th
                          key={person.key}
                          className="px-3 py-3 text-center font-medium"
                        >
                          <span className="block">{person.subject}</span>
                          <span className="text-[11px] font-normal text-[#718078]">
                            {person.department}
                          </span>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {activeReviewers.map((reviewer) => (
                      <tr
                        key={reviewer.id}
                        className="border-t border-[#e5ece8]"
                      >
                        <td className="px-3 py-3">
                          <b>{reviewer.name}</b>
                          <span className="ml-2 text-xs text-[#718078]">
                            {reviewer.department}
                          </span>
                        </td>
                        {people.map((person) => {
                          const checked =
                            assignmentMap[reviewer.id]?.includes(person.key) ||
                            false;
                          return (
                            <td
                              key={person.key}
                              className="px-3 py-3 text-center"
                            >
                              <Checkbox
                                checked={checked}
                                onCheckedChange={(value) =>
                                  toggleAssignment(
                                    reviewer.id,
                                    person.key,
                                    value === true,
                                  )
                                }
                                aria-label={`分配${reviewer.name}评价${person.subject}`}
                              />
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
        <DialogFooter>
          <Button variant="outline" onClick={() => onOpenChange(false)}>
            取消
          </Button>
          <Button onClick={save} className="bg-[#123c32]">
            创建并分配
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

function ReviewDialog({
  open,
  onOpenChange,
  active,
  data,
  setData,
  setNotice,
}: {
  open: boolean;
  onOpenChange: (x: boolean) => void;
  active: { taskId: string; assignmentId: string } | null;
  data: AppData;
  setData: React.Dispatch<React.SetStateAction<AppData>>;
  setNotice: (s: string) => void;
}) {
  const task = data.tasks.find((t) => t.id === active?.taskId);
  const assignment = task?.assignments.find(
    (a) => a.id === active?.assignmentId,
  );
  const tpl = data.templates.find((t) => t.id === task?.templateId);
  const [scores, setScores] = useState<Record<string, number>>({});
  const [invalidScoreIds, setInvalidScoreIds] = useState<string[]>([]);
  const [comment, setComment] = useState('');
  const [remarks, setRemarks] = useState('');
  const [fields, setFields] = useState<Record<string, string>>({});
  useEffect(() => {
    setScores(assignment?.scores || {});
    setInvalidScoreIds([]);
    setComment(assignment?.comment || '');
    setRemarks(assignment?.remarks || '');
    setFields(assignment?.fields || {});
  }, [assignment?.id, assignment?.status]);
  if (!task || !assignment) return null;
  if (!tpl) {
    return (
      <Dialog open={open} onOpenChange={onOpenChange}>
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle>评价表暂时无法打开</DialogTitle>
            <DialogDescription>
              该任务关联的评价模板不存在，请联系管理员恢复模板后重试。
            </DialogDescription>
          </DialogHeader>
          <DialogFooter>
            <Button variant="outline" onClick={() => onOpenChange(false)}>
              关闭
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    );
  }
  const score = weightedScore(
    { ...assignment, status: 'submitted', scores },
    tpl,
  );
  const submit = () => {
    const missing = tpl.criteria.filter(
      (c) =>
        scores[c.id] === undefined || !Number.isFinite(Number(scores[c.id])),
    );
    const outOfRange = tpl.criteria.filter(
      (c) =>
        scores[c.id] !== undefined &&
        Number.isFinite(Number(scores[c.id])) &&
        (Number(scores[c.id]) < 0 || Number(scores[c.id]) > c.max),
    );
    const invalid = [...missing, ...outOfRange];
    if (invalid.length) {
      setInvalidScoreIds(invalid.map((c) => c.id));
      if (missing.length) {
        setNotice(
          `请填写评分：${missing
            .map((c) => `${c.name}（0 至 ${c.max} 分）`)
            .join('、')}`,
        );
      } else {
        const first = outOfRange[0];
        setNotice(
          `${first.name}须填写 0 至 ${first.max} 分，当前为 ${scores[first.id]} 分`,
        );
      }
      return;
    }
    setInvalidScoreIds([]);
    setData((d) => ({
      ...d,
      tasks: d.tasks.map((t) =>
        t.id !== task.id
          ? t
          : {
              ...t,
              assignments: t.assignments.map((a) =>
                a.id !== assignment.id
                  ? a
                  : {
                      ...a,
                      scores,
                      comment,
                      remarks,
                      fields,
                      status: 'submitted',
                      submittedAt: today(),
                    },
              ),
            },
      ),
    }));
    onOpenChange(false);
    setNotice('评价已提交，汇总结果已更新');
  };
  if (tpl.sheetLayout) {
    return (
      <Dialog open={open} onOpenChange={onOpenChange}>
        <DialogContent className="max-h-[96vh] overflow-y-auto p-3 sm:max-w-[min(98vw,1500px)] md:p-5">
          <DialogHeader className="sr-only">
            <DialogTitle>{tpl.name}</DialogTitle>
            <DialogDescription>{task.name}</DialogDescription>
          </DialogHeader>
          <ImportedSheetScoreTable
            task={task}
            assignment={assignment}
            template={tpl}
            scores={scores}
            setScores={setScores}
            comment={comment}
            setComment={setComment}
            remarks={remarks}
            setRemarks={setRemarks}
            fields={fields}
            setFields={setFields}
            score={score}
            invalidScoreIds={invalidScoreIds}
            setInvalidScoreIds={setInvalidScoreIds}
          />
          <DialogFooter className="sticky bottom-0 z-10 mt-0 border border-t-0 bg-white/95 py-3 backdrop-blur">
            <Button variant="outline" onClick={() => onOpenChange(false)}>
              关闭
            </Button>
            {assignment.status !== 'submitted' && (
              <Button onClick={submit} className="bg-[#123c32]">
                <Check />
                提交评价
              </Button>
            )}
          </DialogFooter>
        </DialogContent>
      </Dialog>
    );
  }
  if (tpl.id === 'tpl-stage-report') {
    return (
      <Dialog open={open} onOpenChange={onOpenChange}>
        <DialogContent className="max-h-[94vh] overflow-y-auto p-3 sm:max-w-[min(96vw,1380px)] md:p-5">
          <DialogHeader className="sr-only">
            <DialogTitle>阶段汇报评价表</DialogTitle>
            <DialogDescription>{task.name}</DialogDescription>
          </DialogHeader>
          <StageReportScoreTable
            task={task}
            assignment={assignment}
            template={tpl}
            scores={scores}
            setScores={setScores}
            comment={comment}
            setComment={setComment}
            remarks={remarks}
            setRemarks={setRemarks}
            score={score}
            invalidScoreIds={invalidScoreIds}
            setInvalidScoreIds={setInvalidScoreIds}
          />
          <DialogFooter className="sticky bottom-0 z-10 mt-0 border border-t-0 bg-white/95 backdrop-blur">
            <Button variant="outline" onClick={() => onOpenChange(false)}>
              关闭
            </Button>
            {assignment.status !== 'submitted' && (
              <Button onClick={submit} className="bg-[#123c32]">
                <Check />
                提交评价
              </Button>
            )}
          </DialogFooter>
        </DialogContent>
      </Dialog>
    );
  }
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-h-[92vh] overflow-y-auto sm:max-w-2xl">
        <DialogHeader>
          <DialogTitle>
            {assignment.status === 'submitted'
              ? '查看评价结果'
              : `评价 ${assignment.subject}`}
          </DialogTitle>
          <DialogDescription>
            {task.name} · 评价人：{assignment.reviewer} · 模板：{tpl.name}
          </DialogDescription>
        </DialogHeader>
        <div className="space-y-3">
          {tpl.criteria.map((c, i) => (
            <label
              key={c.id}
              className="block rounded-xl border border-[#e1e9e5] p-4"
            >
              <div className="flex items-start justify-between gap-4">
                <div>
                  <p className="text-sm font-semibold">
                    <span className="mr-2 text-[#8b9891]">
                      {String(i + 1).padStart(2, '0')}
                    </span>
                    {c.name}
                  </p>
                  <p className="mt-1 text-xs text-[#718078]">
                    {c.dimension} · 权重 {c.weight}%{' '}
                    {c.description && `· ${c.description}`}
                  </p>
                </div>
                <div className="flex shrink-0 items-center gap-2">
                  <Input
                    disabled={assignment.status === 'submitted'}
                    className={`h-10 w-20 text-right text-base font-semibold ${
                      invalidScoreIds.includes(c.id)
                        ? 'border-red-500 ring-2 ring-red-100'
                        : ''
                    }`}
                    type="number"
                    min={0}
                    max={c.max}
                    required
                    aria-invalid={invalidScoreIds.includes(c.id)}
                    value={scores[c.id] ?? ''}
                    onChange={(e) => {
                      const rawValue = e.target.value;
                      setScores((s) => {
                        const next = { ...s };
                        if (rawValue === '') delete next[c.id];
                        else next[c.id] = Number(rawValue);
                        return next;
                      });
                      setInvalidScoreIds((ids) =>
                        ids.filter((id) => id !== c.id),
                      );
                    }}
                  />
                  <span className="text-xs text-[#718078]">/ {c.max}</span>
                </div>
              </div>
            </label>
          ))}
          <label className="block text-sm font-medium">
            综合评语
            <Textarea
              disabled={assignment.status === 'submitted'}
              className="mt-1 min-h-24"
              value={comment}
              onChange={(e) => setComment(e.target.value)}
              placeholder="请填写具体表现、优点和改进建议"
            />
          </label>
          <div className="flex items-center justify-between rounded-xl bg-[#eaf5f0] px-4 py-3">
            <span className="text-sm text-[#416257]">当前折算综合分</span>
            <b className="text-xl text-[#123c32]">{score.toFixed(1)}</b>
          </div>
        </div>
        <DialogFooter>
          <Button variant="outline" onClick={() => onOpenChange(false)}>
            关闭
          </Button>
          {assignment.status !== 'submitted' && (
            <Button onClick={submit} className="bg-[#123c32]">
              <Check />
              提交评价
            </Button>
          )}
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

function ImportedSheetScoreTable({
  task,
  assignment,
  template,
  scores,
  setScores,
  comment,
  setComment,
  remarks,
  setRemarks,
  fields,
  setFields,
  score,
  invalidScoreIds,
  setInvalidScoreIds,
}: {
  task: Task;
  assignment: Assignment;
  template: Template;
  scores: Record<string, number>;
  setScores: React.Dispatch<React.SetStateAction<Record<string, number>>>;
  comment: string;
  setComment: (value: string) => void;
  remarks: string;
  setRemarks: (value: string) => void;
  fields: Record<string, string>;
  setFields: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  score: number;
  invalidScoreIds: string[];
  setInvalidScoreIds: React.Dispatch<React.SetStateAction<string[]>>;
}) {
  const layout = template.sheetLayout!;
  const disabled = assignment.status === 'submitted';
  const cells = new Map(
    layout.cells.map((cell) => [`${cell.row}:${cell.col}`, cell]),
  );
  const dynamicText = (cell: SheetCell) => {
    if (cell.role === 'subject')
      return `${cell.text.replace(/[：:].*$/, '')}：${assignment.subject}`;
    if (cell.role === 'reviewer')
      return `${cell.text.replace(/[：:].*$/, '')}：${assignment.reviewer}`;
    if (cell.role === 'stage')
      return `${cell.text.replace(/[：:].*$/, '')}：${formatStageRange(task)}`;
    if (cell.role === 'signature')
      return `${cell.text.replace(/[：:].*$/, '')}：${assignment.reviewer} / ${
        assignment.submittedAt || '提交后自动签署'
      }`;
    return cell.text;
  };
  const renderCell = (cell: SheetCell) => {
    if (cell.role === 'score' && cell.criterionId) {
      const criterion = template.criteria.find(
        (item) => item.id === cell.criterionId,
      );
      if (!criterion) return null;
      return (
        <div className="flex items-center justify-center gap-2 px-2">
          <Input
            disabled={disabled}
            type="number"
            min={0}
            max={criterion.max}
            required
            aria-label={`${criterion.name}评分`}
            aria-invalid={invalidScoreIds.includes(criterion.id)}
            value={scores[criterion.id] ?? ''}
            onChange={(event) => {
              const rawValue = event.target.value;
              setScores((current) => {
                const next = { ...current };
                if (rawValue === '') delete next[criterion.id];
                else next[criterion.id] = Number(rawValue);
                return next;
              });
              setInvalidScoreIds((ids) =>
                ids.filter((id) => id !== criterion.id),
              );
            }}
            className={`h-11 min-w-20 max-w-28 bg-white text-center text-lg font-semibold ${
              invalidScoreIds.includes(criterion.id)
                ? 'border-red-500 ring-2 ring-red-100'
                : 'border-[#59645f]'
            }`}
          />
          <span className="whitespace-nowrap text-xs text-[#56635d]">
            / {criterion.max}
          </span>
        </div>
      );
    }
    if (cell.role === 'comment')
      return (
        <Textarea
          disabled={disabled}
          value={comment}
          onChange={(event) => setComment(event.target.value)}
          className="h-full min-h-20 resize-none border-0 bg-transparent shadow-none"
          placeholder="填写综合评价"
        />
      );
    if (cell.role === 'remarks')
      return (
        <Textarea
          disabled={disabled}
          value={remarks}
          onChange={(event) => setRemarks(event.target.value)}
          className="h-full min-h-28 resize-none border-0 bg-transparent shadow-none"
          placeholder="填写备注"
        />
      );
    if (cell.role === 'input' && cell.fieldId)
      return (
        <Input
          disabled={disabled}
          value={fields[cell.fieldId] || ''}
          onChange={(event) =>
            setFields((current) => ({
              ...current,
              [cell.fieldId!]: event.target.value,
            }))
          }
          className="h-full min-h-10 border-0 bg-transparent shadow-none"
          placeholder={cell.text || '请填写'}
          aria-label={cell.text || '自定义填写项'}
        />
      );
    if (cell.role === 'total')
      return <b className="text-2xl text-[#123c32]">{score.toFixed(1)}</b>;
    return dynamicText(cell);
  };
  return (
    <div className="overflow-x-auto rounded-sm border border-[#2f3834] bg-white">
      <table
        className="table-fixed border-collapse text-[14px] text-[#17231f]"
        style={{
          width: layout.columnWidths.reduce((sum, width) => sum + width, 0),
        }}
      >
        <colgroup>
          {layout.columnWidths.map((width, index) => (
            <col key={index} style={{ width }} />
          ))}
        </colgroup>
        <tbody>
          {Array.from({ length: layout.rowCount }, (_, row) => (
            <tr key={row} style={{ height: layout.rowHeights[row] }}>
              {Array.from({ length: layout.colCount }, (_, col) => {
                const cell = cells.get(`${row}:${col}`);
                if (!cell) return null;
                return (
                  <td
                    key={col}
                    rowSpan={cell.rowSpan}
                    colSpan={cell.colSpan}
                    className="border border-[#2f3834] px-3 py-2"
                    style={{
                      fontWeight: cell.style?.bold ? 700 : undefined,
                      fontSize: cell.style?.fontSize,
                      textAlign: cell.style?.align || 'left',
                      verticalAlign: cell.style?.vertical || 'middle',
                      backgroundColor: cell.style?.background,
                      color: cell.style?.color,
                      whiteSpace: cell.style?.wrap ? 'pre-wrap' : 'normal',
                      overflowWrap: 'anywhere',
                    }}
                  >
                    {renderCell(cell)}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

const stagePointDetails: Record<string, Array<[string, string]>> = {
  'stage-work-result': [
    ['目标完成情况', '工作指标完成情况、进度和质量；'],
    [
      '试用期工作成果',
      '工作成果有数据可展示或具体案例说明，以及解决了什么问题；',
    ],
    ['工作规划', '下一阶段的工作计划、工作任务、节点和重点等安排。'],
  ],
  'stage-key-ability': [
    [
      '沟通协作',
      '善于沟通，能够团结同事，并密切配合同事完成工作任务，善于协调团队内外部关系，将项目实施工作有效推进并达成目标。',
    ],
    [
      '学习能力',
      '具有自我学习提升的意识并落地执行，自驱性强，具有一定的学习力；',
    ],
    ['执行力', '按时按质按量完成个人工作目标；'],
    [
      '问题解决与推动能力',
      '以结果为导向，能预见到问题，同时具备解决问题的决策力和执行力。',
    ],
  ],
  'stage-defense': [
    ['准确性', '回答直击要点；'],
    ['条理性', '论点及论证逻辑清晰，结构合理，回答内容关联性强；'],
    ['答辩效果', '思路敏捷、语言组织、归纳总结能力强、现场反应效果较好。'],
  ],
};

const stageRanges = [
  { label: '好', values: ['44～60', '26～30', '9～10'] },
  { label: '中', values: ['21～43', '16～25', '6～8'] },
  { label: '差', values: ['1～20', '1～15', '1～5'] },
];

function StageReportScoreTable({
  task,
  assignment,
  template,
  scores,
  setScores,
  comment,
  setComment,
  remarks,
  setRemarks,
  score,
  invalidScoreIds,
  setInvalidScoreIds,
}: {
  task: Task;
  assignment: Assignment;
  template: Template;
  scores: Record<string, number>;
  setScores: React.Dispatch<React.SetStateAction<Record<string, number>>>;
  comment: string;
  setComment: (value: string) => void;
  remarks: string;
  setRemarks: (value: string) => void;
  score: number;
  invalidScoreIds: string[];
  setInvalidScoreIds: React.Dispatch<React.SetStateAction<string[]>>;
}) {
  const disabled = assignment.status === 'submitted';
  const cell = 'border border-[#2f3834] px-3 py-3 align-middle';
  const head = `${cell} bg-[#f8faf9] text-center font-semibold`;
  return (
    <div className="overflow-x-auto rounded-sm border border-[#2f3834] bg-white">
      <table className="w-full min-w-[1120px] table-fixed border-collapse text-[14px] text-[#17231f]">
        <colgroup>
          <col className="w-[86px]" />
          <col className="w-[92px]" />
          <col className="w-[30%]" />
          <col className="w-[34%]" />
          <col className="w-[30%]" />
          <col className="w-[150px]" />
        </colgroup>
        <tbody>
          <tr>
            <th
              colSpan={6}
              className="border border-[#2f3834] px-4 py-4 text-center text-[30px] font-medium tracking-[0.12em]"
            >
              阶段汇报评价表
            </th>
          </tr>
          <tr className="text-base">
            <td colSpan={2} className={cell}>
              汇报人：<b>{assignment.subject}</b>
            </td>
            <td colSpan={2} className={cell}>
              评价人：<b>{assignment.reviewer}</b>
            </td>
            <td colSpan={2} className={cell}>
              考察期阶段：<b>{formatStageRange(task)}</b>
            </td>
          </tr>
          <tr>
            <th colSpan={2} className={head}>
              考核评分要素
            </th>
            {template.criteria.map((c) => (
              <th key={c.id} className={head}>
                {c.name}
              </th>
            ))}
            <th rowSpan={2} className={head}>
              备注
            </th>
          </tr>
          <tr>
            <th colSpan={2} className={head}>
              分　值
            </th>
            {template.criteria.map((c) => (
              <th key={c.id} className={`${head} text-lg`}>
                {c.max}
              </th>
            ))}
          </tr>
          <tr>
            <th colSpan={2} className={`${head} text-lg tracking-[0.22em]`}>
              评分要点
            </th>
            {template.criteria.map((c) => (
              <td
                key={c.id}
                className={`${cell} h-[270px] align-middle text-[15px] leading-7`}
              >
                {(stagePointDetails[c.id] || []).map(([title, detail]) => (
                  <p key={title}>
                    <b>{title}：</b>
                    {detail}
                  </p>
                ))}
              </td>
            ))}
            <td className={`${cell} p-2`}>
              <Textarea
                disabled={disabled}
                value={remarks}
                onChange={(e) => setRemarks(e.target.value)}
                className="h-[250px] resize-none border-0 bg-transparent p-2 text-sm shadow-none focus-visible:ring-1"
                placeholder="填写备注"
              />
            </td>
          </tr>
          {stageRanges.map((range, index) => (
            <tr key={range.label}>
              {index === 0 && (
                <th
                  rowSpan={3}
                  className={`${head} text-base tracking-[0.18em]`}
                >
                  评分等级
                </th>
              )}
              <th className={`${head} text-base`}>{range.label}</th>
              {range.values.map((value, valueIndex) => (
                <td
                  key={valueIndex}
                  className={`${cell} py-4 text-center text-base`}
                >
                  {value}
                </td>
              ))}
              <td className={`${cell} bg-[#fafcfb]`} />
            </tr>
          ))}
          <tr>
            <th colSpan={2} className={`${head} text-base`}>
              合计
            </th>
            {template.criteria.map((c) => (
              <td key={c.id} className={`${cell} py-3 text-center`}>
                <div className="inline-flex items-center gap-2">
                  <Input
                    disabled={disabled}
                    type="number"
                    min={0}
                    max={c.max}
                    required
                    aria-invalid={invalidScoreIds.includes(c.id)}
                    value={scores[c.id] ?? ''}
                    onChange={(e) => {
                      const rawValue = e.target.value;
                      setScores((current) => {
                        const next = { ...current };
                        if (rawValue === '') delete next[c.id];
                        else next[c.id] = Number(rawValue);
                        return next;
                      });
                      setInvalidScoreIds((ids) =>
                        ids.filter((id) => id !== c.id),
                      );
                    }}
                    className={`h-11 w-24 bg-white text-center text-lg font-semibold ${
                      invalidScoreIds.includes(c.id)
                        ? 'border-red-500 ring-2 ring-red-100'
                        : 'border-[#7b8882]'
                    }`}
                    aria-label={`${c.name}评分`}
                  />
                  <span className="text-xs text-[#66736d]">/ {c.max}</span>
                </div>
              </td>
            ))}
            <td className={`${cell} bg-[#fafcfb]`} />
          </tr>
          <tr>
            <th colSpan={2} className={`${head} text-base`}>
              综合得分
            </th>
            <td colSpan={4} className={`${cell} py-4`}>
              <div className="flex items-center justify-center gap-3">
                <b className="text-3xl text-[#123c32]">{score.toFixed(1)}</b>
                <span className="text-sm text-[#66736d]">/ 100 分</span>
              </div>
            </td>
          </tr>
          <tr>
            <th colSpan={2} className={`${head} text-base`}>
              综合评价
            </th>
            <td colSpan={4} className={`${cell} p-2`}>
              <Textarea
                disabled={disabled}
                value={comment}
                onChange={(e) => setComment(e.target.value)}
                className="min-h-[110px] resize-none border-0 bg-transparent p-3 text-base shadow-none focus-visible:ring-1"
                placeholder="请填写综合评价、优点和改进建议"
              />
            </td>
          </tr>
          <tr>
            <td colSpan={6} className={`${cell} py-5 text-center text-base`}>
              评价者签名/日期：
              <b className="ml-3">
                {disabled
                  ? `${assignment.reviewer} / ${assignment.submittedAt || today()}`
                  : `${assignment.reviewer} / 提交后自动签署`}
              </b>
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}
