
// ======================================================================
// src/generators/DocumentGenerator.js
// ======================================================================

const TemplateEngine = require('../core/TemplateEngine');
const PizZip = require('pizzip');

class DocumentGenerator {
    constructor() {
        this.templateEngine = new TemplateEngine();
    }

    async generateDocument(templateBuffer, data, options = {}) {
    try {
        // Enhanced validation
        await this.validateDocxBuffer(templateBuffer);
        
        // Validate input data
        this.validateGenerationData(data);

        // Pre-process data for better compatibility
        const processedData = this.preprocessData(data, options);

        // Generate document using template engine
        let generatedBuffer = await this.templateEngine.processTemplate(templateBuffer, processedData);

        // Post-process if needed
        if (options.postProcess) {
            generatedBuffer = await this.postProcessDocument(generatedBuffer, options);
        }

        return generatedBuffer;
    } catch (error) {
        throw new Error(`Document generation failed: ${error.message}`);
    }
}



async validateDocxBuffer(buffer) {
    console.log('🔍 Validating DOCX buffer...');
    console.log(`Buffer size: ${buffer.length} bytes`);
    
    // Check minimum file size (empty DOCX is ~3KB)
    if (buffer.length < 1000) {
        throw new Error(`DOCX file too small (${buffer.length} bytes). Minimum size should be ~3KB`);
    }
    
    // Check maximum file size (100MB)
    if (buffer.length > 100 * 1024 * 1024) {
        throw new Error(`DOCX file too large (${buffer.length} bytes). Maximum size is 100MB`);
    }
    
    // Check ZIP signature (DOCX files are ZIP archives)
    const zipSignature1 = buffer.slice(0, 4);
    const zipSignature2 = buffer.slice(0, 2);
    
    // ZIP file signatures
    const validSignatures = [
        [0x50, 0x4B, 0x03, 0x04], // Standard ZIP
        [0x50, 0x4B, 0x05, 0x06], // Empty ZIP
        [0x50, 0x4B, 0x07, 0x08]  // ZIP with data descriptor
    ];
    
    const hasValidSignature = validSignatures.some(sig => 
        sig.every((byte, i) => buffer[i] === byte)
    );
    
    if (!hasValidSignature) {
        console.error('❌ Invalid ZIP signature:', Array.from(zipSignature1).map(b => '0x' + b.toString(16).padStart(2, '0')));
        throw new Error('Invalid DOCX file: Not a valid ZIP archive. Please ensure the file is a properly saved .docx file');
    }
    
    // Try to read the ZIP structure
    try {
        const PizZip = require('pizzip');
        const zip = new PizZip(buffer);
        
        // Check for essential DOCX files
        const requiredFiles = [
            'word/document.xml',
            '[Content_Types].xml',
            '_rels/.rels'
        ];
        
        const missingFiles = requiredFiles.filter(file => !zip.files[file]);
        
        if (missingFiles.length > 0) {
            throw new Error(`Invalid DOCX structure. Missing required files: ${missingFiles.join(', ')}`);
        }
        
        // Try to read document.xml to ensure it's not corrupted
        const documentXml = zip.files['word/document.xml'].asText();
        if (!documentXml || documentXml.length < 100) {
            throw new Error('Invalid or empty document.xml in DOCX file');
        }
        
        console.log('✅ DOCX file validation passed');
        console.log(`Found ${Object.keys(zip.files).length} files in ZIP archive`);
        
    } catch (error) {
        if (error.message.includes('Invalid DOCX structure') || error.message.includes('Invalid or empty')) {
            throw error;
        }
        console.error('❌ ZIP parsing error:', error.message);
        throw new Error(`Corrupted DOCX file: ${error.message}`);
    }
}


    async generateMultipleDocuments(templateBuffer, dataArray, options = {}) {
        const results = [];
        
        for (let i = 0; i < dataArray.length; i++) {
            try {
                const generatedDoc = await this.generateDocument(templateBuffer, dataArray[i], {
                    ...options,
                    index: i
                });
                
                results.push({
                    index: i,
                    success: true,
                    document: generatedDoc,
                    size: generatedDoc.length
                });
            } catch (error) {
                results.push({
                    index: i,
                    success: false,
                    error: error.message
                });
            }
        }

        return results;
    }

    validateGenerationData(data) {
        if (!data || typeof data !== 'object') {
            throw new Error('Data must be a valid object');
        }

        // Check for circular references
        try {
            JSON.stringify(data);
        } catch (error) {
            if (error.message.includes('circular')) {
                throw new Error('Data contains circular references');
            }
            throw error;
        }

        // Validate data size (prevent memory issues)
        const dataSize = JSON.stringify(data).length;
        if (dataSize > 10 * 1024 * 1024) { // 10MB limit
            throw new Error('Data payload too large (exceeds 10MB)');
        }
    }

    preprocessData(data, options) {
        // Deep clone to avoid mutations
        let processedData = JSON.parse(JSON.stringify(data));

        // Add metadata
        processedData._meta = {
            generatedAt: new Date().toISOString(),
            templateEngine: 'DOCX Template Engine v1.0',
            ...options.metadata
        };

        // Add helper functions
        processedData._helpers = {
            formatDate: (date, format) => {
                const moment = require('moment');
                return moment(date).format(format);
            },
            formatCurrency: (amount, currency = 'USD') => {
                return new Intl.NumberFormat('en-US', {
                    style: 'currency',
                    currency: currency
                }).format(amount);
            },
            formatNumber: (num, decimals = 2) => {
                return Number(num).toFixed(decimals);
            }
        };

        // Normalize text fields (handle Unicode issues)
        processedData = this.normalizeTextFields(processedData);

        return processedData;
    }

    normalizeTextFields(obj) {
        if (typeof obj === 'string') {
            return obj
                .replace(/[""]/g, '"') // Replace smart quotes
                .replace(/['']/g, "'") // Replace smart apostrophes
                .replace(/[–—]/g, '-') // Replace en/em dashes
                .replace(/\u00A0/g, ' ') // Replace non-breaking spaces
                .trim();
        }

        if (Array.isArray(obj)) {
            return obj.map(item => this.normalizeTextFields(item));
        }

        if (obj && typeof obj === 'object') {
            const normalized = {};
            for (const [key, value] of Object.entries(obj)) {
                normalized[key] = this.normalizeTextFields(value);
            }
            return normalized;
        }

        return obj;
    }

    async detectAdvancedFeatures(templateBuffer) {
        try {
            const zip = new PizZip(templateBuffer);
            const documentXml = zip.files['word/document.xml'].asText();
            
            // Check for advanced template features
            const hasConditions = /\$\{#if\s+/.test(documentXml);
            const hasLoops = /\$\{#each\s+/.test(documentXml);
            const hasAggregations = /\$\{[^}]*\|(sum|count|avg|max|min)/.test(documentXml);
            const hasComplexFormatting = /\$\{[^}]*\|(bold|italic|size|color)/.test(documentXml);
            
            return hasConditions || hasLoops || hasAggregations || hasComplexFormatting;
        } catch (error) {
            console.warn('Could not detect advanced features, using basic processing:', error.message);
            return false;
        }
    }

    async postProcessDocument(documentBuffer, options) {
        // Implement post-processing features
        let processedBuffer = documentBuffer;

        // Add page numbers if requested
        if (options.addPageNumbers) {
            processedBuffer = await this.addPageNumbers(processedBuffer);
        }

        // Add headers/footers if requested
        if (options.header || options.footer) {
            processedBuffer = await this.addHeaderFooter(processedBuffer, options);
        }

        // Convert to PDF if requested
        if (options.outputFormat === 'pdf') {
            processedBuffer = await this.convertToPdf(processedBuffer);
        }

        return processedBuffer;
    }

    async addPageNumbers(documentBuffer) {
        // Implement page number addition
        // This would require more complex DOCX manipulation
        console.warn('Page numbers feature not yet implemented');
        return documentBuffer;
    }

    async addHeaderFooter(documentBuffer, options) {
        // Implement header/footer addition
        console.warn('Header/footer feature not yet implemented');
        return documentBuffer;
    }

    async convertToPdf(documentBuffer) {
        // This would require a library like puppeteer or libre-office
        console.warn('PDF conversion feature not yet implemented');
        return documentBuffer;
    }

    async batchGenerate(templateBuffer, dataArray, options = {}) {
        const batchSize = options.batchSize || 10;
        const results = [];

        for (let i = 0; i < dataArray.length; i += batchSize) {
            const batch = dataArray.slice(i, i + batchSize);
            const batchResults = await this.generateMultipleDocuments(templateBuffer, batch, options);
            results.push(...batchResults);

            // Optional delay between batches to prevent system overload
            if (options.batchDelay && i + batchSize < dataArray.length) {
                await this.sleep(options.batchDelay);
            }
        }

        return results;
    }

    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

module.exports = DocumentGenerator;