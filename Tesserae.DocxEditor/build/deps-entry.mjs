/*
 * Vendor bundle entry — collects the pure-JS document libraries used by the docx-editor
 * core and exposes them on window.docxeditordeps. Built as a standalone IIFE that must load
 * BEFORE docx-editor.js (whose vendor imports are shimmed to read from this global).
 *
 * dompurify is intentionally NOT here — Tesserae already provides window.DOMPurify.
 */
import JSZip from 'jszip';
import PizZip from 'pizzip';
import * as xmljs from 'xml-js';
import Docxtemplater from 'docxtemplater';

globalThis.docxeditordeps = { JSZip, PizZip, xmljs, Docxtemplater };
