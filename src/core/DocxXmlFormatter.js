// src/core/DocxXmlFormatter.js

class DocxXmlFormatter {
    constructor() {
        this.formatTagMap = {
            bold: '<w:b/>',
            italic: '<w:i/>',
            underline: '<w:u w:val="single"/>',
            size: (size) => `<w:sz w:val="${size * 2}"/><w:szCs w:val="${size * 2}"/>`, // Word uses half-points
            color: (color) => `<w:color w:val="${this.normalizeColor(color)}"/>`
        };
    }

    /**
     * Apply DOCX formatting to text content
     * @param {string} text - The text content to format
     * @param {Object} formatting - Formatting properties object
     * @returns {string} - Formatted DOCX XML
     */
    applyDocxFormatting(text, formatting) {
        if (!formatting || Object.keys(formatting).length === 0) {
            // No formatting - return plain text wrapped in basic run
            return this.wrapInRun(this.escapeXml(text));
        }

        // Build run properties XML
        const runProperties = this.buildRunProperties(formatting);
        const escapedText = this.escapeXml(text);
        
        // Return formatted run
        return `<w:r>${runProperties}<w:t>${escapedText}</w:t></w:r>`;
    }

    /**
     * Build DOCX run properties XML from formatting object
     * @param {Object} formatting - Formatting properties
     * @returns {string} - DOCX run properties XML
     */
    buildRunProperties(formatting) {
        const properties = [];

        // Apply each formatting property
        for (const [key, value] of Object.entries(formatting)) {
            const formatTag = this.getFormatTag(key, value);
            if (formatTag) {
                properties.push(formatTag);
            }
        }

        // Return wrapped in run properties if we have any
        return properties.length > 0 ? `<w:rPr>${properties.join('')}</w:rPr>` : '';
    }

    /**
     * Get the appropriate DOCX XML tag for a formatting property
     * @param {string} property - Formatting property name
     * @param {*} value - Formatting property value
     * @returns {string} - DOCX XML tag
     */
    getFormatTag(property, value) {
        const tagGenerator = this.formatTagMap[property];
        
        if (typeof tagGenerator === 'function') {
            return tagGenerator(value);
        } else if (typeof tagGenerator === 'string') {
            return tagGenerator;
        }
        
        console.warn(`Unknown formatting property: ${property}`);
        return '';
    }

    /**
     * Wrap plain text in a basic DOCX run
     * @param {string} text - Text to wrap
     * @returns {string} - DOCX run XML
     */
    wrapInRun(text) {
        return `<w:r><w:t>${text}</w:t></w:r>`;
    }

    /**
     * Normalize color values for DOCX
     * @param {string} color - Color value (named colors, hex, etc.)
     * @returns {string} - DOCX-compatible color value
     */
    normalizeColor(color) {
        // Handle common named colors
        const colorMap = {
            'black': '000000',
            'white': 'FFFFFF',
            'red': 'FF0000',
            'green': '00FF00',
            'blue': '0000FF',
            'yellow': 'FFFF00',
            'cyan': '00FFFF',
            'magenta': 'FF00FF',
            'orange': 'FFA500',
            'purple': '800080',
            'brown': 'A52A2A',
            'gray': '808080',
            'grey': '808080'
        };

        const lowerColor = color.toLowerCase();
        
        // Check if it's a named color
        if (colorMap[lowerColor]) {
            return colorMap[lowerColor];
        }
        
        // Check if it's already a hex color (with or without #)
        if (/^#?[0-9A-Fa-f]{6}$/.test(color)) {
            return color.replace('#', '').toUpperCase();
        }
        
        // Default to black if unrecognized
        console.warn(`Unrecognized color: ${color}, defaulting to black`);
        return '000000';
    }

    /**
     * Escape XML special characters
     * @param {string} text - Text to escape
     * @returns {string} - Escaped text
     */
    escapeXml(text) {
        return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    /**
     * Check if a text fragment needs special XML handling
     * @param {string} text - Text to check
     * @returns {boolean} - True if special handling needed
     */
    needsSpecialHandling(text) {
        // Check for line breaks, tabs, or other special characters
        return /[\n\r\t]/.test(text);
    }

    /**
     * Handle special text content like line breaks
     * @param {string} text - Text with special characters
     * @param {Object} formatting - Formatting to apply
     * @returns {string} - Formatted DOCX XML with special handling
     */
    handleSpecialContent(text, formatting = {}) {
        // Handle line breaks
        if (text.includes('\n')) {
            const lines = text.split('\n');
            return lines.map((line, index) => {
                const formattedLine = this.applyDocxFormatting(line, formatting);
                // Add line break after each line except the last
                return index < lines.length - 1 ? 
                    formattedLine + '<w:br/>' : 
                    formattedLine;
            }).join('');
        }
        
        // Handle tabs
        if (text.includes('\t')) {
            text = text.replace(/\t/g, '<w:tab/>');
        }
        
        return this.applyDocxFormatting(text, formatting);
    }
}

module.exports = DocxXmlFormatter;