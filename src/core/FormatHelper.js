// src/core/FormatHelper.js
const moment = require('moment');

class FormatHelper {
    constructor() {
        this.formatters = new Map();
        this.setupDefaultFormatters();
    }

    setupDefaultFormatters() {
        // Text formatters
        this.formatters.set('upper', (value) => String(value).toUpperCase());
        this.formatters.set('lower', (value) => String(value).toLowerCase());
        this.formatters.set('capitalize', (value) => {
            const str = String(value);
            return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
        });
        this.formatters.set('trim', (value) => String(value).trim());
        
        // Number formatters
        this.formatters.set('currency', (value, options = {}) => {
            const num = Number(value) || 0;
            const currency = options.currency || 'USD';
            const locale = options.locale || 'en-US';
            
            return new Intl.NumberFormat(locale, {
                style: 'currency',
                currency: currency
            }).format(num);
        });
        
        this.formatters.set('number', (value, options = {}) => {
            const num = Number(value) || 0;
            const decimals = options.decimals !== undefined ? options.decimals : 2;
            return num.toFixed(decimals);
        });
        
        this.formatters.set('percent', (value) => {
            const num = Number(value) || 0;
            return (num * 100).toFixed(2) + '%';
        });
        
        this.formatters.set('round', (value, places = 0) => {
            const num = Number(value) || 0;
            return Math.round(num * Math.pow(10, places)) / Math.pow(10, places);
        });
        
        // Date formatters
        this.formatters.set('date', (value, format = 'YYYY-MM-DD') => {
            if (!value) return '';
            const date = moment(value);
            return date.isValid() ? date.format(format) : String(value);
        });
        
        this.formatters.set('dateTime', (value, format = 'YYYY-MM-DD HH:mm:ss') => {
            if (!value) return '';
            const date = moment(value);
            return date.isValid() ? date.format(format) : String(value);
        });
        
        this.formatters.set('fromNow', (value) => {
            if (!value) return '';
            const date = moment(value);
            return date.isValid() ? date.fromNow() : String(value);
        });
        
        // Array/Collection formatters
        this.formatters.set('join', (value, separator = ', ') => {
            if (Array.isArray(value)) {
                return value.join(separator);
            }
            return String(value);
        });
        
        this.formatters.set('length', (value) => {
            if (Array.isArray(value) || typeof value === 'string') {
                return value.length;
            }
            return 0;
        });
        
        // Aggregation formatters
        this.formatters.set('sum', (array, field) => {
            if (!Array.isArray(array)) return 0;
            return array.reduce((sum, item) => {
                const val = field ? this.getNestedValue(item, field) : item;
                return sum + (Number(val) || 0);
            }, 0);
        });
        
        this.formatters.set('count', (array) => {
            return Array.isArray(array) ? array.length : 0;
        });
        
        this.formatters.set('avg', (array, field) => {
            if (!Array.isArray(array) || array.length === 0) return 0;
            const sum = this.formatters.get('sum')(array, field);
            return sum / array.length;
        });
        
        // String formatters with length limits
        this.formatters.set('truncate', (value, length = 50) => {
            const str = String(value);
            return str.length > length ? str.substring(0, length) + '...' : str;
        });
        
        // Conditional formatters
        this.formatters.set('default', (value, defaultValue = '') => {
            return (value === null || value === undefined || value === '') ? defaultValue : value;
        });
        
        // XML/HTML formatters
        this.formatters.set('escape', (value) => {
            return String(value)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&apos;');
        });
        
        // Formatting control (these will be handled by DOCX processor)
        this.formatters.set('bold', (value) => ({ value, format: { bold: true } }));
        this.formatters.set('italic', (value) => ({ value, format: { italic: true } }));
        this.formatters.set('underline', (value) => ({ value, format: { underline: true } }));
        this.formatters.set('size', (value, size) => ({ value, format: { size: parseInt(size) } }));
        this.formatters.set('color', (value, color) => ({ value, format: { color: color } }));
    }

    applyFormatters(value, formatters) {
        let result = value;
        let formatting = {};
        
        for (const formatterExpr of formatters) {
            const [name, ...args] = formatterExpr.split(':');
            const formatter = this.formatters.get(name.trim());
            
            if (formatter) {
                try {
                    const formattedResult = formatter(result, ...args.map(arg => arg.trim()));
                    
                    // Handle formatting objects
                    if (formattedResult && typeof formattedResult === 'object' && formattedResult.format) {
                        result = formattedResult.value;
                        formatting = { ...formatting, ...formattedResult.format };
                    } else {
                        result = formattedResult;
                    }
                } catch (error) {
                    console.warn(`Formatter error for "${name}":`, error.message);
                }
            } else {
                console.warn(`Unknown formatter: ${name}`);
            }
        }
        
        // Return formatted result with styling info
        return Object.keys(formatting).length > 0 ? { value: result, formatting } : result;
    }
    
    registerFormatter(name, formatter) {
        this.formatters.set(name, formatter);
    }
    
    getNestedValue(obj, path) {
        return path.split('.').reduce((current, key) => {
            return current && current[key] !== undefined ? current[key] : null;
        }, obj);
    }
}

module.exports = FormatHelper;