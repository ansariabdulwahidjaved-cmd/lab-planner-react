import React, { useState, useEffect, useRef, useMemo } from 'react';
import Chart from 'chart.js/auto';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import * as XLSX from 'xlsx';

const TUTORS = ["Shruti ma'am", "Ravi sir", "AW sir"];
const DAYS_IN_WEEK = 6; // Monday to Saturday only (Sunday removed)

function getMonday(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  return new Date(d.setDate(diff));
}

function formatDateKey(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function getWeekDates(mondayDate) {
  return Array.from({ length: DAYS_IN_WEEK }, (_, i) => {
    const d = new Date(mondayDate);
    d.setDate(mondayDate.getDate() + i);
    return d;
  });
}

function formatWeekRangeText(mondayDate) {
  const dates = getWeekDates(mondayDate);
  const firstDay = dates[0];
  const lastDay = dates[dates.length - 1];
  const opts = { month: 'short', day: 'numeric' };
  return `${firstDay.toLocaleDateString(undefined, opts)} – ${lastDay.toLocaleDateString(undefined, { ...opts, year: 'numeric' })}`;
}

function createInitialState() {
  const mon = getMonday(new Date());
  const dKey = (offset) => {
    const d = new Date(mon);
    d.setDate(mon.getDate() + offset);
    return formatDateKey(d);
  };

  return {
    timings: [
      '09:00 AM - 10:00 AM',
      '10:00 AM - 11:00 AM',
      '11:15 AM - 12:15 PM',
      '01:15 PM - 02:15 PM'
    ],
    batches: [
      '5A(1-21)',
      '5A(22-42)',
      '5B(1-20)',
      '5B(21-40)',
      '6C(1-25)'
    ],
    schedule: {
      [`${dKey(0)}_0`]: [{ batchName: '5A(1-21)', tutor: "Shruti ma'am", topic: "HTML5 Semantic Elements & Forms" }],
      [`${dKey(0)}_1`]: [{ batchName: '5B(1-20)', tutor: "Ravi sir", topic: "JavaScript ES6 Arrow Functions" }],
      [`${dKey(1)}_0`]: [{ batchName: '5A(22-42)', tutor: "AW sir", topic: "CSS Grid & Responsive Flexbox" }],
      [`${dKey(1)}_2`]: [{ batchName: '6C(1-25)', tutor: "Shruti ma'am", topic: "SQL Joins & Normalization" }],
      [`${dKey(2)}_1`]: [{ batchName: '5B(21-40)', tutor: "Ravi sir", topic: "DOM Manipulation & Event Listeners" }],
      [`${dKey(3)}_0`]: [{ batchName: '5A(1-21)', tutor: "AW sir", topic: "REST APIs & Fetch Async/Await" }],
      [`${dKey(5)}_2`]: [{ batchName: '5B(1-20)', tutor: "Shruti ma'am", topic: "React Components & Props Lab" }]
    }
  };
}

export default function App() {
  const [appState, setAppState] = useState(() => {
    try {
      const saved = localStorage.getItem('pinkGreenLabPlannerState_v1');
      return saved ? JSON.parse(saved) : createInitialState();
    } catch {
      return createInitialState();
    }
  });

  const [currentMonday, setCurrentMonday] = useState(() => getMonday(new Date()));
  const [viewMode, setViewMode] = useState('mode-both');
  const [selectedBatchForTouch, setSelectedBatchForTouch] = useState(null);
  const [dragOverCell, setDragOverCell] = useState(null);

  // Inputs for new timing and batch
  const [newTimeInput, setNewTimeInput] = useState('');
  const [newBatchInput, setNewBatchInput] = useState('');

  // Dashboard Filters
  const [dashboardFilters, setDashboardFilters] = useState({
    period: 'weekly',
    tutor: 'ALL',
    batch: 'ALL'
  });

  // Lecture Modal State
  const [lectureModal, setLectureModal] = useState({
    isOpen: false,
    cellKey: null,
    index: null,
    batchName: '',
    dateLabel: '',
    timeLabel: '',
    tutor: '',
    topic: ''
  });

  // Copy Week Modal State
  const [copyModal, setCopyModal] = useState({
    isOpen: false,
    sourceDate: '',
    keepTutors: true,
    keepTopics: false,
    overwriteTarget: true
  });

  const [isExportingPdf, setIsExportingPdf] = useState(false);

  // Refs for Charts & PDF Capture
  const batchChartCanvasRef = useRef(null);
  const tutorClassChartCanvasRef = useRef(null);
  const dateTimelineChartCanvasRef = useRef(null);
  const pdfReportContentRef = useRef(null);
  const timelineBoardRef = useRef(null);
  const topicInputRef = useRef(null);

  const chartInstancesRef = useRef({
    batchChart: null,
    tutorClassChart: null,
    dateTimelineChart: null
  });

  // Save state to localStorage
  useEffect(() => {
    localStorage.setItem('pinkGreenLabPlannerState_v1', JSON.stringify(appState));
  }, [appState]);

  const weekDates = useMemo(() => getWeekDates(currentMonday), [currentMonday]);
  const todayKey = formatDateKey(new Date());

  // Filtered Lectures Calculation
  const filteredLectures = useMemo(() => {
    const today = new Date();
    const tKey = formatDateKey(today);
    const weekSet = new Set(weekDates.map(formatDateKey));
    const refMonth = currentMonday.getMonth();
    const refYear = currentMonday.getFullYear();

    const results = [];

    Object.keys(appState.schedule).forEach(cellKey => {
      const lastUnderscore = cellKey.lastIndexOf('_');
      const dateStr = cellKey.substring(0, lastUnderscore);
      const tIdx = Number(cellKey.substring(lastUnderscore + 1));
      const [y, m, d] = dateStr.split('-').map(Number);
      const dateObj = new Date(y, m - 1, d);

      if (dateObj.getDay() === 0) return;

      if (dashboardFilters.period === 'today' && dateStr !== tKey) return;
      if (dashboardFilters.period === 'weekly' && !weekSet.has(dateStr)) return;
      if (
        dashboardFilters.period === 'monthly' &&
        (dateObj.getMonth() !== refMonth || dateObj.getFullYear() !== refYear)
      ) {
        return;
      }

      (appState.schedule[cellKey] || []).forEach((lec, idx) => {
        if (dashboardFilters.tutor !== 'ALL' && lec.tutor !== dashboardFilters.tutor) return;
        if (dashboardFilters.batch !== 'ALL' && lec.batchName !== dashboardFilters.batch) return;

        results.push({
          cellKey,
          index: idx,
          dateStr,
          dateObj,
          timeLabel: appState.timings[tIdx] || 'Slot',
          batchName: lec.batchName,
          tutor: lec.tutor || 'Unassigned',
          topic: lec.topic || ''
        });
      });
    });

    return results.sort((a, b) => a.dateStr.localeCompare(b.dateStr));
  }, [appState.schedule, appState.timings, currentMonday, dashboardFilters, weekDates]);

  // KPI Metrics Calculation
  const kpiMetrics = useMemo(() => {
    const totalLectures = filteredLectures.length;
    const assignedCount = filteredLectures.filter(r => r.tutor && r.tutor !== 'Unassigned').length;
    const tutorPct = totalLectures ? Math.round((assignedCount / totalLectures) * 100) : 0;

    const topicsCount = filteredLectures.filter(r => r.topic && r.topic.trim().length > 0).length;
    const topicsPct = totalLectures ? Math.round((topicsCount / totalLectures) * 100) : 0;

    const uniqueBatches = new Set(filteredLectures.map(r => r.batchName)).size;
    const batchCoveragePct = appState.batches.length
      ? Math.round((uniqueBatches / appState.batches.length) * 100)
      : 0;

    return {
      totalLectures,
      assignedCount,
      tutorPct: Math.max(0, Math.min(100, tutorPct)),
      topicsCount,
      topicsPct: Math.max(0, Math.min(100, topicsPct)),
      uniqueBatches,
      batchCoveragePct: Math.max(0, Math.min(100, batchCoveragePct))
    };
  }, [filteredLectures, appState.batches.length]);

  // Render Chart.js Visuals
  useEffect(() => {
    const pinkMintPalette = ['#f472b6', '#6ee7b7', '#db2777', '#34d399', '#f9a8d4', '#10b981'];
    const batchLabels = appState.batches;
    const batchCounts = batchLabels.map(b => filteredLectures.filter(r => r.batchName === b).length);

    if (chartInstancesRef.current.batchChart) {
      chartInstancesRef.current.batchChart.destroy();
    }
    if (batchChartCanvasRef.current) {
      chartInstancesRef.current.batchChart = new Chart(batchChartCanvasRef.current, {
        type: 'doughnut',
        data: {
          labels: batchLabels,
          datasets: [{
            data: batchCounts,
            backgroundColor: pinkMintPalette,
            borderColor: '#ffffff',
            borderWidth: 2
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          animation: false,
          plugins: {
            legend: {
              position: 'right',
              labels: { font: { family: 'Plus Jakarta Sans', weight: '600' } }
            }
          }
        }
      });
    }

    const tutorColors = {
      "Shruti ma'am": '#ec4899',
      "Ravi sir": '#10b981',
      "AW sir": '#f472b6'
    };

    const tutorDatasets = TUTORS.map(tutorName => ({
      label: tutorName,
      backgroundColor: tutorColors[tutorName],
      borderRadius: 5,
      data: batchLabels.map(b => filteredLectures.filter(r => r.batchName === b && r.tutor === tutorName).length)
    }));

    if (chartInstancesRef.current.tutorClassChart) {
      chartInstancesRef.current.tutorClassChart.destroy();
    }
    if (tutorClassChartCanvasRef.current) {
      chartInstancesRef.current.tutorClassChart = new Chart(tutorClassChartCanvasRef.current, {
        type: 'bar',
        data: {
          labels: batchLabels,
          datasets: tutorDatasets
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          animation: false,
          scales: {
            y: { beginAtZero: true, ticks: { stepSize: 1 } }
          },
          plugins: {
            legend: { position: 'top' }
          }
        }
      });
    }

    const uniqueDates = [...new Set(filteredLectures.map(r => r.dateStr))];
    const formattedDates = uniqueDates.map(dStr => {
      const [y, m, d] = dStr.split('-').map(Number);
      return new Date(y, m - 1, d).toLocaleDateString(undefined, { weekday: 'short', day: 'numeric', month: 'short' });
    });

    const dateDatasets = TUTORS.map(tutorName => ({
      label: tutorName,
      backgroundColor: tutorColors[tutorName],
      borderRadius: 4,
      data: uniqueDates.map(dStr => filteredLectures.filter(r => r.dateStr === dStr && r.tutor === tutorName).length)
    }));

    if (chartInstancesRef.current.dateTimelineChart) {
      chartInstancesRef.current.dateTimelineChart.destroy();
    }
    if (dateTimelineChartCanvasRef.current) {
      chartInstancesRef.current.dateTimelineChart = new Chart(dateTimelineChartCanvasRef.current, {
        type: 'bar',
        data: {
          labels: formattedDates.length ? formattedDates : ['No Lectures in Filter'],
          datasets: dateDatasets
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          animation: false,
          scales: {
            x: { stacked: true },
            y: { stacked: true, beginAtZero: true, ticks: { stepSize: 1 } }
          }
        }
      });
    }

    return () => {
      if (chartInstancesRef.current.batchChart) chartInstancesRef.current.batchChart.destroy();
      if (chartInstancesRef.current.tutorClassChart) chartInstancesRef.current.tutorClassChart.destroy();
      if (chartInstancesRef.current.dateTimelineChart) chartInstancesRef.current.dateTimelineChart.destroy();
    };
  }, [filteredLectures, appState.batches, viewMode]);

  // Actions: Add / Delete Timings & Batches
  const handleAddTiming = () => {
    const val = newTimeInput.trim();
    if (!val) return;
    setAppState(prev => ({
      ...prev,
      timings: [...prev.timings, val]
    }));
    setNewTimeInput('');
  };

  const handleDeleteTiming = (timeIndex) => {
    setAppState(prev => {
      const newTimings = prev.timings.filter((_, idx) => idx !== timeIndex);
      const updatedSchedule = {};
      Object.keys(prev.schedule).forEach(key => {
        const lastUnderscore = key.lastIndexOf('_');
        const datePart = key.substring(0, lastUnderscore);
        const tIdx = Number(key.substring(lastUnderscore + 1));
        if (tIdx < timeIndex) {
          updatedSchedule[key] = prev.schedule[key];
        } else if (tIdx > timeIndex) {
          updatedSchedule[`${datePart}_${tIdx - 1}`] = prev.schedule[key];
        }
      });
      return {
        ...prev,
        timings: newTimings,
        schedule: updatedSchedule
      };
    });
  };

  const handleAddBatch = () => {
    const val = newBatchInput.trim();
    if (!val) return;
    setAppState(prev => ({
      ...prev,
      batches: [...prev.batches, val]
    }));
    setNewBatchInput('');
  };

  const handleDeleteBatch = (batchIndex) => {
    const removed = appState.batches[batchIndex];
    setAppState(prev => {
      const newBatches = prev.batches.filter((_, idx) => idx !== batchIndex);
      const updatedSchedule = {};
      Object.keys(prev.schedule).forEach(key => {
        updatedSchedule[key] = prev.schedule[key].filter(item => item.batchName !== removed);
      });
      return {
        ...prev,
        batches: newBatches,
        schedule: updatedSchedule
      };
    });
    if (dashboardFilters.batch === removed) {
      setDashboardFilters(prev => ({ ...prev, batch: 'ALL' }));
    }
  };

  // Drag & Drop / Cell Assignment
  const assignBatchToCell = (targetCell, batchName, sourceCell = null, sourceIndex = null) => {
    setAppState(prev => {
      const nextSchedule = { ...prev.schedule };
      let movedLecture = null;

      if (sourceCell && sourceCell !== targetCell && nextSchedule[sourceCell]) {
        const sourceList = [...nextSchedule[sourceCell]];
        movedLecture = sourceList.splice(sourceIndex, 1)[0];
        nextSchedule[sourceCell] = sourceList;
      } else if (sourceCell === targetCell) {
        return prev;
      }

      const targetList = nextSchedule[targetCell] ? [...nextSchedule[targetCell]] : [];
      const exists = targetList.some(item => item.batchName === batchName);
      if (!exists) {
        targetList.push(
          movedLecture || {
            batchName,
            tutor: '',
            topic: ''
          }
        );
      }
      nextSchedule[targetCell] = targetList;
      return { ...prev, schedule: nextSchedule };
    });
  };

  const handleDropOnCell = (e, targetCell) => {
    e.preventDefault();
    setDragOverCell(null);
    const rawData = e.dataTransfer.getData('text/plain');
    if (!rawData) return;
    try {
      const { batchName, sourceCell, sourceIndex } = JSON.parse(rawData);
      assignBatchToCell(targetCell, batchName, sourceCell, sourceIndex);
    } catch (err) {
      console.error(err);
    }
  };

  const handleRemoveBatchFromCell = (e, cellKey, index) => {
    e.stopPropagation();
    setAppState(prev => {
      const nextSchedule = { ...prev.schedule };
      if (nextSchedule[cellKey]) {
        const list = [...nextSchedule[cellKey]];
        list.splice(index, 1);
        nextSchedule[cellKey] = list;
      }
      return { ...prev, schedule: nextSchedule };
    });
  };

  // Lecture Modal Handlers
  const openLectureModal = (cellKey, index, dateObj, timeLabel) => {
    const lecture = appState.schedule[cellKey]?.[index];
    if (!lecture) return;

    setLectureModal({
      isOpen: true,
      cellKey,
      index,
      batchName: lecture.batchName,
      dateLabel: `📅 ${dateObj.toLocaleDateString(undefined, { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' })}`,
      timeLabel: `⏰ ${timeLabel}`,
      tutor: lecture.tutor || '',
      topic: lecture.topic || ''
    });

    setTimeout(() => {
      topicInputRef.current?.focus();
    }, 50);
  };

  const closeLectureModal = () => {
    setLectureModal(prev => ({ ...prev, isOpen: false }));
  };

  const handleSaveLectureDetails = () => {
    const { cellKey, index, tutor, topic } = lectureModal;
    if (!cellKey || index === null) return;

    setAppState(prev => {
      const nextSchedule = { ...prev.schedule };
      const list = [...(nextSchedule[cellKey] || [])];
      if (list[index]) {
        list[index] = {
          ...list[index],
          tutor,
          topic: topic.trim()
        };
        nextSchedule[cellKey] = list;
      }
      return { ...prev, schedule: nextSchedule };
    });

    closeLectureModal();
  };

  // Copy Week Modal Logic
  const openCopyWeekModal = () => {
    const prevMon = new Date(currentMonday);
    prevMon.setDate(currentMonday.getDate() - 7);
    setCopyModal({
      isOpen: true,
      sourceDate: formatDateKey(prevMon),
      keepTutors: true,
      keepTopics: false,
      overwriteTarget: true
    });
  };

  const closeCopyWeekModal = () => {
    setCopyModal(prev => ({ ...prev, isOpen: false }));
  };

  const copySourcePreviewText = useMemo(() => {
    if (!copyModal.sourceDate) return '';
    const [y, m, d] = copyModal.sourceDate.split('-').map(Number);
    const sourceMon = getMonday(new Date(y, m - 1, d));
    const sourceDates = getWeekDates(sourceMon);
    let count = 0;
    sourceDates.forEach(dObj => {
      const dKey = formatDateKey(dObj);
      appState.timings.forEach((_, tIdx) => {
        count += (appState.schedule[`${dKey}_${tIdx}`] || []).length;
      });
    });
    return `Source Week (Mon–Sat): ${formatWeekRangeText(sourceMon)} — ${count} scheduled lecture(s) found`;
  }, [copyModal.sourceDate, appState.schedule, appState.timings]);

  const handleConfirmCopyWeek = () => {
    if (!copyModal.sourceDate) {
      alert('Please choose a date from the source week you want to copy.');
      return;
    }

    const [y, m, d] = copyModal.sourceDate.split('-').map(Number);
    const sourceMonday = getMonday(new Date(y, m - 1, d));

    if (formatDateKey(sourceMonday) === formatDateKey(currentMonday)) {
      alert('Source week and Current Target week are the same. Please select a different week to copy from.');
      return;
    }

    const sourceDates = getWeekDates(sourceMonday);
    const targetDates = getWeekDates(currentMonday);
    let copiedCount = 0;

    setAppState(prev => {
      const nextSchedule = { ...prev.schedule };

      for (let dayOffset = 0; dayOffset < DAYS_IN_WEEK; dayOffset++) {
        const srcDateKey = formatDateKey(sourceDates[dayOffset]);
        const tgtDateKey = formatDateKey(targetDates[dayOffset]);

        prev.timings.forEach((_, tIdx) => {
          const srcCellKey = `${srcDateKey}_${tIdx}`;
          const tgtCellKey = `${tgtDateKey}_${tIdx}`;

          if (copyModal.overwriteTarget) {
            delete nextSchedule[tgtCellKey];
          }

          const srcLectures = prev.schedule[srcCellKey] || [];
          if (srcLectures.length > 0) {
            const targetList = nextSchedule[tgtCellKey] ? [...nextSchedule[tgtCellKey]] : [];

            srcLectures.forEach(lec => {
              const alreadyExists = targetList.some(item => item.batchName === lec.batchName);
              if (!alreadyExists) {
                targetList.push({
                  batchName: lec.batchName,
                  tutor: copyModal.keepTutors ? (lec.tutor || '') : '',
                  topic: copyModal.keepTopics ? (lec.topic || '') : ''
                });
                copiedCount++;
              }
            });

            nextSchedule[tgtCellKey] = targetList;
          }
        });
      }

      return { ...prev, schedule: nextSchedule };
    });

    closeCopyWeekModal();
    setTimeout(() => {
      alert(`Successfully copied ${copiedCount} lecture(s) from ${formatWeekRangeText(sourceMonday)} to ${formatWeekRangeText(currentMonday)}.`);
    }, 50);
  };

  // Clear Current Week
  const handleClearWeek = () => {
    if (window.confirm('Clear all scheduled lectures for this Monday–Saturday week?')) {
      const weekDateKeys = weekDates.map(formatDateKey);
      setAppState(prev => {
        const nextSchedule = { ...prev.schedule };
        Object.keys(nextSchedule).forEach(key => {
          const datePart = key.substring(0, key.lastIndexOf('_'));
          if (weekDateKeys.includes(datePart)) {
            delete nextSchedule[key];
          }
        });
        return { ...prev, schedule: nextSchedule };
      });
    }
  };

  // Export Multi-Sheet Excel (.xlsx)
  const exportToExcel = () => {
    try {
      const wb = XLSX.utils.book_new();

      // SHEET 1: Current Week Timetable Grid (Mon - Sat)
      const gridHeader = [
        'Timings \ Mon–Sat',
        ...weekDates.map(d => d.toLocaleDateString(undefined, { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' }))
      ];
      const gridRows = [gridHeader];

      appState.timings.forEach((timeLabel, tIdx) => {
        const row = [timeLabel];
        weekDates.forEach(dateObj => {
          const cellKey = `${formatDateKey(dateObj)}_${tIdx}`;
          const cellLectures = appState.schedule[cellKey] || [];
          if (cellLectures.length === 0) {
            row.push('');
          } else {
            const formattedCell = cellLectures.map(lec => {
              const parts = [`Batch: ${lec.batchName}`];
              if (lec.tutor) parts.push(`Tutor: ${lec.tutor}`);
              if (lec.topic) parts.push(`Topic: ${lec.topic}`);
              return parts.join(' | ');
            }).join('\n');
            row.push(formattedCell);
          }
        });
        gridRows.push(row);
      });

      const wsGrid = XLSX.utils.aoa_to_sheet(gridRows);
      wsGrid['!cols'] = [{ wch: 24 }, ...weekDates.map(() => ({ wch: 32 }))];
      XLSX.utils.book_append_sheet(wb, wsGrid, 'Current Week Grid');

      // SHEET 2: Filtered Lecture Log
      const logRows = [
        ['Date', 'Day', 'Timing Slot', 'Class Batch', 'Tutor Name', 'Topics Covered']
      ];

      filteredLectures.forEach(r => {
        logRows.push([
          r.dateStr,
          r.dateObj.toLocaleDateString(undefined, { weekday: 'long' }),
          r.timeLabel,
          r.batchName,
          r.tutor,
          r.topic || 'Not Logged'
        ]);
      });

      const wsLog = XLSX.utils.aoa_to_sheet(logRows);
      wsLog['!cols'] = [
        { wch: 14 },
        { wch: 14 },
        { wch: 22 },
        { wch: 16 },
        { wch: 18 },
        { wch: 45 }
      ];
      XLSX.utils.book_append_sheet(wb, wsLog, 'Filtered Lecture Log');

      // SHEET 3: Summary Analytics
      const summaryRows = [
        ['LAB PLANNER — ANALYTICS SUMMARY'],
        ['Filter Period', dashboardFilters.period.toUpperCase()],
        ['Tutor Filter', dashboardFilters.tutor],
        ['Batch Filter', dashboardFilters.batch],
        ['Total Lectures in Filter', filteredLectures.length],
        [],
        ['Class Batch', 'Total Lectures', ...TUTORS, 'Unassigned']
      ];

      appState.batches.forEach(batchName => {
        const batchRecs = filteredLectures.filter(r => r.batchName === batchName);
        const tutorCounts = TUTORS.map(tName => batchRecs.filter(r => r.tutor === tName).length);
        const unassignedCount = batchRecs.filter(r => r.tutor === 'Unassigned').length;
        summaryRows.push([
          batchName,
          batchRecs.length,
          ...tutorCounts,
          unassignedCount
        ]);
      });

      const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
      wsSummary['!cols'] = [
        { wch: 20 },
        { wch: 16 },
        { wch: 16 },
        { wch: 16 },
        { wch: 16 },
        { wch: 16 }
      ];
      XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary Analytics');

      const fileDate = formatDateKey(new Date());
      const periodLabel = dashboardFilters.period.toUpperCase();
      XLSX.writeFile(wb, `Lab_Planner_Export_${periodLabel}_${fileDate}.xlsx`);
    } catch (err) {
      console.error('Excel Export Error:', err);
      alert('Could not export Excel file. Please check your browser console.');
    }
  };

  // Export Graphic Summary to PDF
  const exportDashboardToPDF = async () => {
    const prevMode = viewMode;
    try {
      setIsExportingPdf(true);

      if (viewMode === 'mode-scheduler') {
        setViewMode('mode-both');
        await new Promise(r => setTimeout(r, 250));
      }

      const timelineBoard = timelineBoardRef.current;
      const prevMaxHeight = timelineBoard ? timelineBoard.style.maxHeight : '';
      const prevOverflow = timelineBoard ? timelineBoard.style.overflow : '';
      if (timelineBoard) {
        timelineBoard.style.maxHeight = 'none';
        timelineBoard.style.overflow = 'visible';
      }

      const captureElement = pdfReportContentRef.current;
      const canvas = await html2canvas(captureElement, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#fff5f9'
      });

      if (timelineBoard) {
        timelineBoard.style.maxHeight = prevMaxHeight;
        timelineBoard.style.overflow = prevOverflow;
      }
      if (prevMode === 'mode-scheduler') {
        setViewMode(prevMode);
      }

      const pdf = new jsPDF('p', 'mm', 'a4');
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();

      pdf.setFillColor(252, 231, 243);
      pdf.rect(0, 0, pageWidth, 28, 'F');
      pdf.setDrawColor(16, 185, 129);
      pdf.setLineWidth(1.4);
      pdf.line(0, 28, pageWidth, 28);

      pdf.setTextColor(131, 24, 67);
      pdf.setFont('helvetica', 'bold');
      pdf.setFontSize(15);
      pdf.text('LAB PLANNER — EXECUTIVE LECTURE & TOPICS REPORT', 12, 12);

      pdf.setTextColor(6, 78, 59);
      pdf.setFont('helvetica', 'normal');
      pdf.setFontSize(9.5);
      const periodLabel = dashboardFilters.period.toUpperCase();
      const tutorLabel = dashboardFilters.tutor === 'ALL' ? 'All Tutors' : dashboardFilters.tutor;
      const batchLabel = dashboardFilters.batch === 'ALL' ? 'All Batches' : dashboardFilters.batch;
      const generatedOn = new Date().toLocaleString();
      pdf.text(
        `Period: ${periodLabel}   |   Tutor: ${tutorLabel}   |   Batch: ${batchLabel}   |   Generated: ${generatedOn}`,
        12,
        21
      );

      const margin = 10;
      const contentWidth = pageWidth - margin * 2;
      const imgHeight = (canvas.height * contentWidth) / canvas.width;
      const imgData = canvas.toDataURL('image/png');

      let startY = 34;
      let heightLeft = imgHeight;
      let position = startY;

      pdf.addImage(imgData, 'PNG', margin, position, contentWidth, imgHeight);
      heightLeft -= (pageHeight - startY);

      while (heightLeft > 0) {
        position = heightLeft - imgHeight + margin;
        pdf.addPage();
        pdf.addImage(imgData, 'PNG', margin, position, contentWidth, imgHeight);
        heightLeft -= (pageHeight - margin * 2);
      }

      const fileDate = formatDateKey(new Date());
      pdf.save(`Lab_Planner_Summary_${periodLabel}_${fileDate}.pdf`);
    } catch (err) {
      console.error('PDF Export Error:', err);
      alert('Could not generate PDF. Please check your browser console.');
    } finally {
      setIsExportingPdf(false);
    }
  };

  return (
    <div className="app-root">
      {/* Light Pink & Pastel Green Header */}
      <header className="app-header">
        <div className="brand-title">
          <div className="brand-crest">LP</div>
          <div>
            <h1>Lab Planner</h1>
            <span className="brand-subtitle">Mon–Sat Timetable, Lecture Topics & Visual Analytics</span>
          </div>
        </div>

        {/* View Mode Switcher */}
        <div className="view-switcher">
          <button
            className={`view-tab ${viewMode === 'mode-both' ? 'active' : ''}`}
            onClick={() => setViewMode('mode-both')}
          >
            🌸🌿 Split View
          </button>
          <button
            className={`view-tab ${viewMode === 'mode-scheduler' ? 'active' : ''}`}
            onClick={() => setViewMode('mode-scheduler')}
          >
            📅 Timetable
          </button>
          <button
            className={`view-tab ${viewMode === 'mode-dashboard' ? 'active' : ''}`}
            onClick={() => setViewMode('mode-dashboard')}
          >
            📊 Graphic Dashboard
          </button>
        </div>

        {/* Calendar Week Navigation, Copy Week, Excel & PDF Export */}
        <div className="calendar-controls">
          <button
            className="btn ghost-duo"
            onClick={() => {
              const prev = new Date(currentMonday);
              prev.setDate(prev.getDate() - 7);
              setCurrentMonday(prev);
            }}
          >
            &larr; Prev
          </button>
          <button
            className="btn pastel-green-btn"
            onClick={() => setCurrentMonday(getMonday(new Date()))}
          >
            Today
          </button>
          <span className="week-label">{formatWeekRangeText(currentMonday)}</span>
          <input
            type="date"
            title="Jump to date"
            value={formatDateKey(weekDates[0])}
            onChange={(e) => {
              if (e.target.value) {
                const [y, m, d] = e.target.value.split('-').map(Number);
                setCurrentMonday(getMonday(new Date(y, m - 1, d)));
              }
            }}
          />
          <button
            className="btn ghost-duo"
            onClick={() => {
              const next = new Date(currentMonday);
              next.setDate(next.getDate() + 7);
              setCurrentMonday(next);
            }}
          >
            Next &rarr;
          </button>
          <button
            className="btn pastel-green-btn"
            title="Copy timetable from a selected week to this current week"
            onClick={openCopyWeekModal}
          >
            📋 Copy Week
          </button>
          <button
            className="btn pastel-green-btn"
            title="Export Timetable & Summary to Excel (.xlsx)"
            onClick={exportToExcel}
          >
            📊 Export Excel
          </button>
          <button
            className="btn pink-primary"
            title="Export Graphic Summary to PDF"
            disabled={isExportingPdf}
            onClick={exportDashboardToPDF}
          >
            {isExportingPdf ? '⏳ Generating PDF...' : '📄 Export PDF'}
          </button>
          <button className="btn danger-outline" onClick={handleClearWeek}>
            Clear Week
          </button>
        </div>
      </header>

      <main id="workspace" className={`workspace ${viewMode}`}>
        {/* 1. TIMETABLE & BATCHES SECTION (MON - SAT) */}
        <section id="scheduler-view" className="scheduler-grid">
          <div className="classic-panel">
            <div className="panel-title-bar">
              <h2>📜 Academic Calendar Grid (Mon – Sat)</h2>
              <div className="timing-controls">
                <input
                  type="text"
                  placeholder="New Timing (e.g., 02:30 PM - 03:30 PM)"
                  value={newTimeInput}
                  onChange={(e) => setNewTimeInput(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && handleAddTiming()}
                />
                <button className="btn pastel-green-btn" onClick={handleAddTiming}>
                  + Add Timing
                </button>
              </div>
            </div>

            <div className="table-wrapper">
              <table id="timetable">
                <thead>
                  <tr>
                    <th className="corner-header">Timings \ Mon–Sat</th>
                    {weekDates.map((dateObj) => {
                      const dateKey = formatDateKey(dateObj);
                      const dayName = dateObj.toLocaleDateString(undefined, { weekday: 'short' });
                      const dateDisplay = dateObj.toLocaleDateString(undefined, { day: 'numeric', month: 'short' });
                      return (
                        <th key={dateKey} className={dateKey === todayKey ? 'today-col' : ''}>
                          <span className="day-name">{dayName}</span>
                          <span className="date-num">{dateDisplay}</span>
                        </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody>
                  {appState.timings.map((time, tIdx) => (
                    <tr key={`${time}_${tIdx}`}>
                      <td className="timing-cell">
                        <div className="header-content">
                          <span>{time}</span>
                          <button
                            className="delete-btn"
                            title="Delete Timing"
                            onClick={() => handleDeleteTiming(tIdx)}
                          >
                            &times;
                          </button>
                        </div>
                      </td>
                      {weekDates.map((dateObj) => {
                        const dateKey = formatDateKey(dateObj);
                        const cellKey = `${dateKey}_${tIdx}`;
                        const cellLectures = appState.schedule[cellKey] || [];
                        const isToday = dateKey === todayKey;
                        const isDragOver = dragOverCell === cellKey;

                        return (
                          <td
                            key={cellKey}
                            className={`drop-cell ${isToday ? 'today-cell' : ''} ${isDragOver ? 'drag-over' : ''}`}
                            onDragOver={(e) => {
                              e.preventDefault();
                              if (dragOverCell !== cellKey) setDragOverCell(cellKey);
                            }}
                            onDragLeave={() => setDragOverCell(null)}
                            onDrop={(e) => handleDropOnCell(e, cellKey)}
                            onClick={() => {
                              if (selectedBatchForTouch) {
                                assignBatchToCell(cellKey, selectedBatchForTouch);
                                setSelectedBatchForTouch(null);
                              }
                            }}
                          >
                            {cellLectures.map((lectureObj, bInCellIdx) => {
                              const hasDetails = Boolean(
                                lectureObj.tutor || (lectureObj.topic && lectureObj.topic.trim())
                              );
                              return (
                                <div
                                  key={`${lectureObj.batchName}_${bInCellIdx}`}
                                  className={`scheduled-batch ${hasDetails ? 'completed' : ''}`}
                                  draggable
                                  onDragStart={(e) => {
                                    e.dataTransfer.setData(
                                      'text/plain',
                                      JSON.stringify({
                                        batchName: lectureObj.batchName,
                                        sourceCell: cellKey,
                                        sourceIndex: bInCellIdx
                                      })
                                    );
                                  }}
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    openLectureModal(cellKey, bInCellIdx, dateObj, time);
                                  }}
                                >
                                  <div className="scheduled-top">
                                    <span>{lectureObj.batchName}</span>
                                    <button
                                      className="delete-btn"
                                      title="Remove lecture"
                                      onClick={(e) => handleRemoveBatchFromCell(e, cellKey, bInCellIdx)}
                                    >
                                      &times;
                                    </button>
                                  </div>
                                  <div className="scheduled-meta">
                                    {lectureObj.tutor && (
                                      <span className="meta-tutor">👩‍🏫 {lectureObj.tutor}</span>
                                    )}
                                    {lectureObj.topic && (
                                      <span className="meta-topic">📖 {lectureObj.topic}</span>
                                    )}
                                    {!hasDetails && (
                                      <span className="click-hint">+ Add Tutor & Topics</span>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Right Sidebar: Batches */}
          <aside className="classic-panel batches-sidebar">
            <h2>🎓 Class Batches</h2>
            <p className="helper-text">
              Drag a batch into any Monday–Saturday slot. <strong>Click any scheduled lecture</strong> to assign a Tutor (
              <em>Shruti ma'am, Ravi sir, AW sir</em>) and <strong>Add Topics</strong> covered.
            </p>

            <div className="batch-controls">
              <input
                type="text"
                placeholder="e.g., 5A(1-21)"
                value={newBatchInput}
                onChange={(e) => setNewBatchInput(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && handleAddBatch()}
              />
              <button className="btn pink-primary" onClick={handleAddBatch}>
                + Add
              </button>
            </div>

            <div className="batches-list">
              {appState.batches.map((batch, bIdx) => (
                <div
                  key={`${batch}_${bIdx}`}
                  className={`batch-item ${selectedBatchForTouch === batch ? 'selected-touch' : ''}`}
                  draggable
                  onDragStart={(e) => {
                    e.dataTransfer.setData(
                      'text/plain',
                      JSON.stringify({ batchName: batch, sourceCell: null })
                    );
                  }}
                  onClick={() => {
                    setSelectedBatchForTouch((prev) => (prev === batch ? null : batch));
                  }}
                >
                  <span>{batch}</span>
                  <button
                    className="delete-btn"
                    title="Delete Batch"
                    onClick={(e) => {
                      e.stopPropagation();
                      handleDeleteBatch(bIdx);
                    }}
                  >
                    &times;
                  </button>
                </div>
              ))}
            </div>
          </aside>
        </section>

        {/* 2. GRAPHIC ANALYTICS DASHBOARD SECTION */}
        <section id="dashboard-view" className="dashboard-section">
          {/* Filter Bar */}
          <div className="filter-bar">
            <div className="filter-title">
              <h2>🏛️ Executive Visual Dashboard</h2>
            </div>

            <div className="filter-controls">
              <div className="period-pills">
                {['today', 'weekly', 'monthly', 'all'].map((periodKey) => (
                  <button
                    key={periodKey}
                    className={`period-btn ${dashboardFilters.period === periodKey ? 'active' : ''}`}
                    onClick={() =>
                      setDashboardFilters((prev) => ({ ...prev, period: periodKey }))
                    }
                  >
                    {periodKey === 'all'
                      ? 'All Time'
                      : periodKey.charAt(0).toUpperCase() + periodKey.slice(1)}
                  </button>
                ))}
              </div>

              <div className="filter-select-group">
                <span>Tutor:</span>
                <select
                  value={dashboardFilters.tutor}
                  onChange={(e) =>
                    setDashboardFilters((prev) => ({ ...prev, tutor: e.target.value }))
                  }
                >
                  <option value="ALL">All Tutors</option>
                  {TUTORS.map((t) => (
                    <option key={t} value={t}>
                      {t}
                    </option>
                  ))}
                </select>
              </div>

              <div className="filter-select-group">
                <span>Batch:</span>
                <select
                  value={dashboardFilters.batch}
                  onChange={(e) =>
                    setDashboardFilters((prev) => ({ ...prev, batch: e.target.value }))
                  }
                >
                  <option value="ALL">All Batches</option>
                  {appState.batches.map((b) => (
                    <option key={b} value={b}>
                      {b}
                    </option>
                  ))}
                </select>
              </div>

              <button className="btn pastel-green-btn" onClick={exportToExcel}>
                📊 Export Excel
              </button>
              <button
                className="btn pink-primary"
                disabled={isExportingPdf}
                onClick={exportDashboardToPDF}
              >
                {isExportingPdf ? '⏳ Generating PDF...' : '📄 Export PDF Report'}
              </button>
            </div>
          </div>

          {/* Printable Graphic Area */}
          <div ref={pdfReportContentRef} id="pdf-report-content" className="pdf-capture-area">
            {/* SVG Radial KPI Gauges */}
            <div className="kpi-grid">
              <div className="kpi-card">
                <svg viewBox="0 0 36 36" className="radial-ring">
                  <path
                    className="ring-bg"
                    d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831"
                  />
                  <path
                    className="ring-meter"
                    stroke="#db2777"
                    strokeDasharray={`${kpiMetrics.totalLectures > 0 ? 100 : 0}, 100`}
                    d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831"
                  />
                  <text x="18" y="19" className="ring-text">
                    {kpiMetrics.totalLectures}
                  </text>
                </svg>
                <div className="kpi-info">
                  <span>Total Lectures</span>
                  <strong>{kpiMetrics.totalLectures}</strong>
                  <span className="kpi-sub">In Selected Period</span>
                </div>
              </div>

              <div className="kpi-card">
                <svg viewBox="0 0 36 36" className="radial-ring">
                  <path
                    className="ring-bg"
                    d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831"
                  />
                  <path
                    className="ring-meter"
                    stroke="#10b981"
                    strokeDasharray={`${kpiMetrics.tutorPct}, 100`}
                    d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831"
                  />
                  <text x="18" y="19" className="ring-text">
                    {kpiMetrics.tutorPct}%
                  </text>
                </svg>
                <div className="kpi-info">
                  <span>Tutor Assigned</span>
                  <strong>
                    {kpiMetrics.assignedCount}/{kpiMetrics.totalLectures}
                  </strong>
                  <span className="kpi-sub">Faculty Coverage</span>
                </div>
              </div>

              <div className="kpi-card">
                <svg viewBox="0 0 36 36" className="radial-ring">
                  <path
                    className="ring-bg"
                    d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831"
                  />
                  <path
                    className="ring-meter"
                    stroke="#f472b6"
                    strokeDasharray={`${kpiMetrics.topicsPct}, 100`}
                    d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831"
                  />
                  <text x="18" y="19" className="ring-text">
                    {kpiMetrics.topicsPct}%
                  </text>
                </svg>
                <div className="kpi-info">
                  <span>Topics Logged</span>
                  <strong>
                    {kpiMetrics.topicsCount}/{kpiMetrics.totalLectures}
                  </strong>
                  <span className="kpi-sub">Syllabus Coverage</span>
                </div>
              </div>

              <div className="kpi-card">
                <svg viewBox="0 0 36 36" className="radial-ring">
                  <path
                    className="ring-bg"
                    d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831"
                  />
                  <path
                    className="ring-meter"
                    stroke="#059669"
                    strokeDasharray={`${kpiMetrics.batchCoveragePct}, 100`}
                    d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831"
                  />
                  <text x="18" y="19" className="ring-text">
                    {kpiMetrics.batchCoveragePct}%
                  </text>
                </svg>
                <div className="kpi-info">
                  <span>Active Batches</span>
                  <strong>
                    {kpiMetrics.uniqueBatches}/{appState.batches.length}
                  </strong>
                  <span className="kpi-sub">Classes Scheduled</span>
                </div>
              </div>
            </div>

            {/* Graphic Charts Grid */}
            <div className="charts-grid">
              <div className="chart-card">
                <div className="chart-header">
                  <h3>1. Lectures Taken Per Batch</h3>
                  <span className="chart-badge">Pink & Mint Doughnut</span>
                </div>
                <div className="canvas-holder">
                  <canvas ref={batchChartCanvasRef}></canvas>
                </div>
              </div>

              <div className="chart-card">
                <div className="chart-header">
                  <h3>2. Tutor Lectures by Class Batch</h3>
                  <span className="chart-badge">Grouped Bar Breakdown</span>
                </div>
                <div className="canvas-holder">
                  <canvas ref={tutorClassChartCanvasRef}></canvas>
                </div>
              </div>

              <div className="chart-card full-width">
                <div className="chart-header">
                  <h3>3. Date-Wise Tutor & Class Lecture Distribution</h3>
                  <span className="chart-badge">Chronological Stacked Visuals</span>
                </div>
                <div className="canvas-holder" style={{ height: '230px' }}>
                  <canvas ref={dateTimelineChartCanvasRef}></canvas>
                </div>
                <div ref={timelineBoardRef} className="timeline-board">
                  {filteredLectures.map((r) => {
                    const initials =
                      r.tutor !== 'Unassigned'
                        ? r.tutor
                            .split(' ')
                            .map((w) => w[0])
                            .join('')
                            .toUpperCase()
                        : '?';

                    return (
                      <div
                        key={`${r.cellKey}_${r.index}`}
                        className="visual-slot-card"
                        onClick={() =>
                          openLectureModal(r.cellKey, r.index, r.dateObj, r.timeLabel)
                        }
                      >
                        <div className="v-card-top">
                          <span className="v-batch-pill">{r.batchName}</span>
                          <span className="v-date-tag">
                            📅{' '}
                            {r.dateObj.toLocaleDateString(undefined, {
                              weekday: 'short',
                              day: 'numeric',
                              month: 'short'
                            })}
                          </span>
                        </div>
                        <div className="v-tutor-bar">
                          <span className="v-tutor-avatar">{initials}</span>
                          <span>{r.tutor}</span>
                          <span
                            style={{
                              marginLeft: 'auto',
                              fontSize: '0.7rem',
                              color: 'var(--text-muted)'
                            }}
                          >
                            {r.timeLabel}
                          </span>
                        </div>
                        <div className="v-topic-box">
                          {r.topic && r.topic.trim() ? (
                            `📖 ${r.topic}`
                          ) : (
                            <span
                              style={{
                                color: 'var(--text-muted)',
                                fontStyle: 'italic'
                              }}
                            >
                              No topic added yet (Click to add)
                            </span>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>
        </section>
      </main>

      {/* Copy Week Timetable Modal */}
      <div className={`modal-overlay ${copyModal.isOpen ? 'active' : ''}`}>
        <div className="modal-card" style={{ maxWidth: '480px' }}>
          <div className="modal-header">
            <h3>📋 Copy Timetable to Current Week</h3>
            <button className="delete-btn" onClick={closeCopyWeekModal}>
              &times;
            </button>
          </div>

          <div className="modal-body">
            <div className="slot-info-bar">
              <span>Target (Current Viewed Week):</span>
              <strong>{formatWeekRangeText(currentMonday)} (Mon–Sat)</strong>
            </div>

            <div className="form-group">
              <label htmlFor="copy-source-date">
                Select any date from the Source Week you want to copy from:
              </label>
              <div style={{ display: 'flex', gap: '0.5rem' }}>
                <input
                  type="date"
                  id="copy-source-date"
                  style={{ flex: 1 }}
                  value={copyModal.sourceDate}
                  onChange={(e) =>
                    setCopyModal((prev) => ({ ...prev, sourceDate: e.target.value }))
                  }
                />
                <button
                  type="button"
                  className="btn pastel-green-btn"
                  onClick={() => {
                    const prevMon = new Date(currentMonday);
                    prevMon.setDate(currentMonday.getDate() - 7);
                    setCopyModal((prev) => ({
                      ...prev,
                      sourceDate: formatDateKey(prevMon)
                    }));
                  }}
                >
                  Use Previous Week
                </button>
              </div>
              <span
                style={{
                  fontSize: '0.8rem',
                  color: 'var(--green-main)',
                  fontWeight: 700,
                  marginTop: '0.25rem'
                }}
              >
                {copySourcePreviewText}
              </span>
            </div>

            <div className="form-group" style={{ gap: '0.6rem', marginTop: '0.2rem' }}>
              <label>Copy Options:</label>
              <label className="checkbox-row">
                <input
                  type="checkbox"
                  checked={copyModal.keepTutors}
                  onChange={(e) =>
                    setCopyModal((prev) => ({ ...prev, keepTutors: e.target.checked }))
                  }
                />
                <span>Keep assigned Tutor names (Shruti ma'am, Ravi sir, AW sir)</span>
              </label>
              <label className="checkbox-row">
                <input
                  type="checkbox"
                  checked={copyModal.keepTopics}
                  onChange={(e) =>
                    setCopyModal((prev) => ({ ...prev, keepTopics: e.target.checked }))
                  }
                />
                <span>Also copy Lecture Topics (leave unchecked for fresh topic entries)</span>
              </label>
              <label className="checkbox-row">
                <input
                  type="checkbox"
                  checked={copyModal.overwriteTarget}
                  onChange={(e) =>
                    setCopyModal((prev) => ({
                      ...prev,
                      overwriteTarget: e.target.checked
                    }))
                  }
                />
                <span>Replace existing lectures in the current week before copying</span>
              </label>
            </div>
          </div>

          <div className="modal-footer">
            <button type="button" className="btn ghost-duo" onClick={closeCopyWeekModal}>
              Cancel
            </button>
            <button
              type="button"
              className="btn pink-primary"
              onClick={handleConfirmCopyWeek}
            >
              📋 Copy to Current Week
            </button>
          </div>
        </div>
      </div>

      {/* Tutor & Add Topics Modal */}
      <div className={`modal-overlay ${lectureModal.isOpen ? 'active' : ''}`}>
        <div className="modal-card">
          <div className="modal-header">
            <h3>Batch: {lectureModal.batchName}</h3>
            <button className="delete-btn" onClick={closeLectureModal}>
              &times;
            </button>
          </div>

          <div className="modal-body">
            <div className="slot-info-bar">
              <span>{lectureModal.dateLabel}</span>
              <span>{lectureModal.timeLabel}</span>
            </div>

            <div className="form-group">
              <label htmlFor="tutor-select">Select Tutor Who Conducted Lecture:</label>
              <select
                id="tutor-select"
                value={lectureModal.tutor}
                onChange={(e) =>
                  setLectureModal((prev) => ({ ...prev, tutor: e.target.value }))
                }
              >
                <option value="">-- Select Tutor --</option>
                {TUTORS.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </select>
            </div>

            <div className="form-group">
              <label htmlFor="topic-input">Add Topics (Textbox for Lecture Topics):</label>
              <textarea
                id="topic-input"
                ref={topicInputRef}
                placeholder="e.g., Arrays & Linked Lists, SQL Joins, Network Topologies..."
                value={lectureModal.topic}
                onChange={(e) =>
                  setLectureModal((prev) => ({ ...prev, topic: e.target.value }))
                }
              />
            </div>
          </div>

          <div className="modal-footer">
            <button type="button" className="btn ghost-duo" onClick={closeLectureModal}>
              Cancel
            </button>
            <button
              type="button"
              className="btn pink-primary"
              onClick={handleSaveLectureDetails}
            >
              Save Lecture Details
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
