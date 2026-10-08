/**
 * excel-importer.js - Smart Excel (.xlsx) Parser for Coach Timesheets
 * Automatically handles CoachTimesheet formats with serial dates, fractional times,
 * descriptions, and remarks.
 */

import * as XLSX from 'xlsx';

export class ExcelTimesheetImporter {
  /**
   * Convert Excel 1900-based serial date to YYYY-MM-DD
   */
  static parseExcelDate(serial) {
    if (!serial) return null;
    
    // If it's already an ISO or date string (e.g. "2026-09-06" or "06/09/2026")
    if (typeof serial === 'string' && serial.includes('-')) {
      return serial.trim();
    }

    const num = Number(serial);
    if (isNaN(num) || num < 1000) return null;

    // Excel epoch 1899-12-30 in UTC (handles 1900 leap year bug)
    const totalDays = Math.floor(num);
    const epoch = new Date(Date.UTC(1899, 11, 30));
    epoch.setUTCDate(epoch.getUTCDate() + totalDays);

    const yyyy = epoch.getUTCFullYear();
    const mm = String(epoch.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(epoch.getUTCDate()).padStart(2, '0');

    return `${yyyy}-${mm}-${dd}`;
  }

  /**
   * Convert Excel fraction of day (e.g. 0.5833333333333334) to 24h & 12h times
   */
  static parseExcelTime(fraction) {
    if (fraction === undefined || fraction === null || fraction === '') {
      return { time24: '12:00', displayTime: '12:00 PM', hours: 12, minutes: 0 };
    }

    // If it's already a string like "14:00" or "2:00 PM"
    if (typeof fraction === 'string' && fraction.includes(':')) {
      const parts = fraction.trim().split(':');
      let h = parseInt(parts[0], 10) || 0;
      let m = parseInt(parts[1], 10) || 0;
      const period = fraction.toLowerCase().includes('pm') || h >= 12 ? 'PM' : 'AM';
      let h12 = h % 12;
      if (h12 === 0) h12 = 12;
      return {
        time24: `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`,
        displayTime: `${String(h12).padStart(2, '0')}:${String(m).padStart(2, '0')} ${period}`,
        hours: h,
        minutes: m
      };
    }

    const num = Number(fraction);
    if (isNaN(num)) {
      return { time24: '12:00', displayTime: '12:00 PM', hours: 12, minutes: 0 };
    }

    // Calculate hours and minutes from day fraction
    const totalSeconds = Math.round(num * 86400);
    const hours = Math.floor(totalSeconds / 3600) % 24;
    const minutes = Math.floor((totalSeconds % 3600) / 60);

    const period = hours >= 12 ? 'PM' : 'AM';
    let h12 = hours % 12;
    if (h12 === 0) h12 = 12;

    return {
      time24: `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`,
      displayTime: `${String(h12).padStart(2, '0')}:${String(minutes).padStart(2, '0')} ${period}`,
      hours,
      minutes
    };
  }

  /**
   * Parse duration strings like "50mins", "1hour", "1.5hour" to integer minutes
   */
  static parseDuration(val) {
    if (!val) return 60;
    if (typeof val === 'number') return Math.round(val);

    const str = String(val).toLowerCase().trim();
    if (str.includes('hour') || str.includes('hr')) {
      const num = parseFloat(str.replace(/[^0-9.]/g, ''));
      return Math.round((num || 1) * 60);
    }
    if (str.includes('min')) {
      const num = parseInt(str.replace(/[^0-9]/g, ''), 10);
      return num || 60;
    }
    const fallback = parseInt(str.replace(/[^0-9]/g, ''), 10);
    return fallback || 60;
  }

  /**
   * Parse raw ArrayBuffer or File from .xlsx file
   */
  static parseWorkbook(arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // Read rows as 2D array
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    // Look for header row containing "Date" and "Description" or "Time"
    let headerIdx = -1;
    let colMap = {
      date: 0,
      day: 1,
      time: 2,
      duration: 3,
      description: 4,
      totalClass: 5,
      remarks: 6
    };

    for (let r = 0; r < Math.min(rows.length, 10); r++) {
      const row = rows[r];
      if (Array.isArray(row)) {
        const rowStr = row.map(c => String(c).toLowerCase().trim());
        const dateCol = rowStr.findIndex(s => s === 'date');
        const descCol = rowStr.findIndex(s => s === 'description' || s === 'class' || s === 'category');
        if (dateCol !== -1 && (descCol !== -1 || rowStr.includes('time'))) {
          headerIdx = r;
          colMap.date = dateCol;
          colMap.day = rowStr.findIndex(s => s === 'day');
          colMap.time = rowStr.findIndex(s => s === 'time');
          colMap.duration = rowStr.findIndex(s => s === 'duration');
          colMap.description = descCol !== -1 ? descCol : 4;
          colMap.totalClass = rowStr.findIndex(s => s === 'total class' || s === 'total');
          colMap.remarks = rowStr.findIndex(s => s === 'remarks' || s === 'note' || s === 'notes');
          break;
        }
      }
    }

    const startRow = headerIdx !== -1 ? headerIdx + 1 : 4;
    const parsedClasses = [];
    const detectedCategories = new Set();

    for (let i = startRow; i < rows.length; i++) {
      const row = rows[i];
      if (!Array.isArray(row) || row.length === 0) continue;

      const rawDate = row[colMap.date];
      if (!rawDate) continue;

      const dateStr = this.parseExcelDate(rawDate);
      if (!dateStr) continue;

      const rawTime = colMap.time !== -1 ? row[colMap.time] : '';
      const timeObj = this.parseExcelTime(rawTime);

      const rawDuration = colMap.duration !== -1 ? row[colMap.duration] : '60';
      const durationMins = this.parseDuration(rawDuration);

      const rawCategory = colMap.description !== -1 ? String(row[colMap.description] || '').trim() : '';
      const rawRemarks = colMap.remarks !== -1 ? String(row[colMap.remarks] || '').trim() : '';

      // For this project, any .xlsx file imported is strictly under the Aerosplash category
      const category = 'Aerosplash';
      detectedCategories.add(category);

      // Preserve description (e.g. LTS, PreComp, Swim Clinic) and remarks in session note
      let combinedNote = '';
      if (rawCategory && rawCategory.toLowerCase() !== 'aerosplash') {
        if (rawRemarks) {
          combinedNote = `${rawCategory} - ${rawRemarks}`;
        } else {
          combinedNote = rawCategory;
        }
      } else {
        combinedNote = rawRemarks;
      }

      // Create timestamp
      const [y, m, d] = dateStr.split('-').map(Number);
      const timestamp = new Date(y, m - 1, d, timeObj.hours, timeObj.minutes).getTime();

      parsedClasses.push({
        id: 'cls_import_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
        date: dateStr,
        time: timeObj.time24,
        displayTime: timeObj.displayTime,
        timestamp,
        duration: durationMins,
        category: 'Aerosplash',
        note: combinedNote,
        importedFrom: firstSheetName,
        createdAt: new Date().toISOString()
      });
    }

    return {
      sheetName: firstSheetName,
      classes: parsedClasses,
      categories: ['Aerosplash'],
      totalCount: parsedClasses.length
    };
  }
}

import { exportCoachTimesheetXLSX, getClassCreditUnits, calculateManagerBonus } from './timesheet-template.js';

export { getClassCreditUnits, calculateManagerBonus, exportCoachTimesheetXLSX };

/**
 * ExcelTimesheetExporter - Generates authentic Coach Timesheet Excel (.xlsx) files
 * Faithfully matches 'JieHern_CoachTimesheet September2026.xlsx' in cell borders,
 * fonts, date/time numFmts, formulas, and summary fee calculations.
 */
export class ExcelTimesheetExporter {
  /**
   * Return a Blob matching the exact formatting, borders, and formulas of the sample timesheet
   */
  static toBlob(classes, options = {}) {
    const opts = typeof options === 'string' ? { sheetName: options } : (options || {});
    return exportCoachTimesheetXLSX(classes, opts);
  }
}


