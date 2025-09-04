const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const morgan = require('morgan');
const path = require('path');
const multer = require('multer');
const { MongoClient } = require('mongodb');
const TemplateManager = require('./managers/TemplateManager');
const DocumentGenerator = require('./generators/DocumentGenerator');
const TemplateValidator = require('./validators/TemplateValidator');

class DocxTemplateServer {
    constructor() {
        this.app = express();
        this.port = process.env.PORT || 3000;
        this.mongoUrl = process.env.MONGO_URL || 'mongodb+srv://prabeshsuwal1234:Kathmandu-123@cluster0.1teqdim.mongodb.net/docx_templates';
        this.db = null;
        
        this.setupMiddleware();
        this.setupRoutes();
    }

    setupMiddleware() {
        // Security and logging
        this.app.use(helmet());
        this.app.use(cors());
        this.app.use(morgan('combined'));
        this.app.use(express.json({ limit: '50mb' }));
        this.app.use(express.urlencoded({ extended: true, limit: '50mb' }));
        
        // Static files
        this.app.use(express.static(path.join(__dirname, 'public')));
        
        // File upload configuration
        const storage = multer.memoryStorage();
        this.upload = multer({ 
            storage,
            fileFilter: (req, file, cb) => {
                if (file.mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
                    cb(null, true);
                } else {
                    cb(new Error('Only .docx files are allowed'), false);
                }
            },
            limits: { fileSize: 10 * 1024 * 1024 } // 10MB limit
        });
    }

    async connectDatabase() {
        try {
            const client = new MongoClient(this.mongoUrl);
            await client.connect();
            this.db = client.db();
            console.log('Connected to MongoDB');
        } catch (error) {
            console.error('MongoDB connection error:', error);
            process.exit(1);
        }
    }

    setupRoutes() {
        // Initialize managers
        const templateManager = new TemplateManager();
        const documentGenerator = new DocumentGenerator();
        const templateValidator = new TemplateValidator();

        // UI Routes
        this.app.get('/', (req, res) => {
            res.sendFile(path.join(__dirname, 'public', 'index.html'));
        });

        // Template Management Routes
        this.app.post('/api/templates', this.upload.single('template'), async (req, res) => {
            try {
                if (!req.file) {
                    return res.status(400).json({ error: 'No template file uploaded' });
                }

                const templateData = {
                    name: req.body.name || req.file.originalname,
                    description: req.body.description || '',
                    author: req.body.author || 'Unknown',
                    buffer: req.file.buffer
                };

                const result = await templateManager.uploadTemplate(this.db, templateData);
                res.json(result);
            } catch (error) {
                console.error('Template upload error:', error);
                res.status(500).json({ error: error.message });
            }
        });

        this.app.get('/api/templates', async (req, res) => {
            try {
                const templates = await templateManager.listTemplates(this.db);
                res.json(templates);
            } catch (error) {
                res.status(500).json({ error: error.message });
            }
        });

        this.app.get('/api/templates/:id', async (req, res) => {
            try {
                const template = await templateManager.getTemplate(this.db, req.params.id);
                if (!template) {
                    return res.status(404).json({ error: 'Template not found' });
                }
                res.json(template);
            } catch (error) {
                res.status(500).json({ error: error.message });
            }
        });

        // Template Validation Routes
        this.app.post('/api/templates/:id/validate', async (req, res) => {
            try {
                const template = await templateManager.getTemplate(this.db, req.params.id);
                if (!template) {
                    return res.status(404).json({ error: 'Template not found' });
                }

                const validation = await templateValidator.validateTemplate(template.buffer);
                res.json(validation);
            } catch (error) {
                res.status(500).json({ error: error.message });
            }
        });

        // Document Generation Routes
        this.app.post('/api/documents/generate', async (req, res) => {
            try {
                const { templateId, data, options = {} } = req.body;
                
                if (!templateId || !data) {
                    return res.status(400).json({ error: 'Template ID and data are required' });
                }

                const template = await templateManager.getTemplate(this.db, templateId);
                if (!template) {
                    return res.status(404).json({ error: 'Template not found' });
                }

                const generatedDoc = await documentGenerator.generateDocument(
                    template.buffer, 
                    data, 
                    options
                );

                // Store generation log
                await this.db.collection('document_logs').insertOne({
                    templateId,
                    generatedAt: new Date(),
                    dataHash: this.hashData(data),
                    success: true
                });

                res.set({
                    'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    'Content-Disposition': `attachment; filename="generated_${Date.now()}.docx"`
                });
                res.send(generatedDoc);
            } catch (error) {
                console.error('Document generation error:', error);
                
                // Log failed generation
                if (req.body.templateId) {
                    await this.db.collection('document_logs').insertOne({
                        templateId: req.body.templateId,
                        generatedAt: new Date(),
                        error: error.message,
                        success: false
                    });
                }
                
                res.status(500).json({ error: error.message });
            }
        });

        // Delete template route
        this.app.delete('/api/templates/:id', async (req, res) => {
            try {
                const result = await templateManager.deleteTemplate(this.db, req.params.id);
                res.json(result);
            } catch (error) {
                res.status(500).json({ error: error.message });
            }
        });

        // Analytics Routes
        this.app.get('/api/analytics/templates', async (req, res) => {
            try {
                const analytics = await this.db.collection('document_logs').aggregate([
                    {
                        $group: {
                            _id: '$templateId',
                            totalGenerations: { $sum: 1 },
                            successfulGenerations: {
                                $sum: { $cond: ['$success', 1, 0] }
                            },
                            lastGenerated: { $max: '$generatedAt' }
                        }
                    }
                ]).toArray();
                
                res.json(analytics);
            } catch (error) {
                res.status(500).json({ error: error.message });
            }
        });


        this.app.post('/api/generate-direct', this.upload.single('template'), async (req, res) => {
    try {
        console.log('Direct generation request received');
        
        // Validate inputs
        if (!req.file) {
            return res.status(400).json({ 
                error: 'No template file uploaded. Please include a .docx file in the "template" field.',
                code: 'NO_FILE'
            });
        }

        // Enhanced file validation
        console.log(`Received file: ${req.file.originalname}`);
        console.log(`File size: ${req.file.size} bytes`);
        console.log(`File mimetype: ${req.file.mimetype}`);
        
        // Check file extension
        const fileExt = req.file.originalname.toLowerCase().split('.').pop();
        if (fileExt !== 'docx') {
            return res.status(400).json({
                error: `Invalid file extension: .${fileExt}. Only .docx files are supported.`,
                code: 'INVALID_EXTENSION'
            });
        }

        // Check MIME type (but be flexible since some systems report different types)
        const validMimeTypes = [
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'application/octet-stream', // Some systems report this
            'application/zip' // DOCX is essentially a ZIP file
        ];
        
        if (!validMimeTypes.includes(req.file.mimetype)) {
            console.warn(`⚠️  Unexpected MIME type: ${req.file.mimetype}, but proceeding...`);
        }

        if (!req.body.data) {
            return res.status(400).json({ 
                error: 'No JSON data provided. Please include JSON data in the "data" field.',
                code: 'NO_DATA'
            });
        }

        // Parse JSON data
        let jsonData;
        try {
            jsonData = typeof req.body.data === 'string' ? 
                JSON.parse(req.body.data) : req.body.data;
        } catch (parseError) {
            return res.status(400).json({ 
                error: 'Invalid JSON data: ' + parseError.message,
                code: 'INVALID_JSON'
            });
        }

        console.log('Processing template:', req.file.originalname);
        console.log('Data keys:', Object.keys(jsonData));

        // Create a debug copy of the uploaded file for inspection
        if (process.env.NODE_ENV === 'development') {
            const fs = require('fs');
            const debugPath = `debug_${Date.now()}_${req.file.originalname}`;
            fs.writeFileSync(debugPath, req.file.buffer);
            console.log(`🔍 Debug: Saved uploaded file to ${debugPath}`);
        }

        // Initialize document generator
        const documentGenerator = new DocumentGenerator();
        
        // Parse options from request
        const options = {
            addPageNumbers: req.body.addPageNumbers === 'true' || req.body.addPageNumbers === true,
            outputFormat: req.body.outputFormat || 'docx',
            metadata: {
                originalFilename: req.file.originalname,
                processedAt: new Date().toISOString(),
                dataKeys: Object.keys(jsonData).length
            }
        };

        // Generate document
        console.log('Generating document...');
        const generatedDoc = await documentGenerator.generateDocument(
            req.file.buffer, 
            jsonData, 
            options
        );

        // Prepare filename
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const baseName = req.file.originalname.replace('.docx', '');
        const filename = `${baseName}_generated_${timestamp}.docx`;

        // Log successful generation
        console.log('Document generated successfully:', filename);
        
        // Optional: Store generation log in database
        if (this.db) {
            try {
                await this.db.collection('generation_logs').insertOne({
                    originalFilename: req.file.originalname,
                    generatedFilename: filename,
                    generatedAt: new Date(),
                    fileSize: generatedDoc.length,
                    originalFileSize: req.file.size,
                    dataKeys: Object.keys(jsonData),
                    success: true,
                    userAgent: req.headers['user-agent'],
                    ip: req.ip || req.connection.remoteAddress
                });
            } catch (logError) {
                console.warn('Failed to log generation:', logError.message);
            }
        }

        // Return generated document
        res.set({
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': `attachment; filename="${filename}"`,
            'Content-Length': generatedDoc.length,
            'X-Generated-At': new Date().toISOString(),
            'X-Original-Template': req.file.originalname,
            'X-Original-Size': req.file.size,
            'X-Generated-Size': generatedDoc.length
        });

        res.send(generatedDoc);

    } catch (error) {
        console.error('Direct generation error:', error);
        
        // Enhanced error categorization
        let errorCode = 'GENERATION_FAILED';
        let statusCode = 500;
        
        if (error.message.includes('Invalid DOCX file') || error.message.includes('Not a valid ZIP')) {
            errorCode = 'INVALID_DOCX_FILE';
            statusCode = 400;
        } else if (error.message.includes('DOCX file too small') || error.message.includes('DOCX file too large')) {
            errorCode = 'INVALID_FILE_SIZE';
            statusCode = 400;
        } else if (error.message.includes('Corrupted DOCX file')) {
            errorCode = 'CORRUPTED_FILE';
            statusCode = 400;
        } else if (error.message.includes('Data contains circular references') || error.message.includes('Data payload too large')) {
            errorCode = 'INVALID_DATA';
            statusCode = 400;
        }
        
        // Log failed generation
        if (this.db) {
            try {
                await this.db.collection('generation_logs').insertOne({
                    originalFilename: req.file ? req.file.originalname : 'unknown',
                    generatedAt: new Date(),
                    success: false,
                    error: error.message,
                    errorCode: errorCode,
                    userAgent: req.headers['user-agent'],
                    ip: req.ip || req.connection.remoteAddress
                });
            } catch (logError) {
                console.warn('Failed to log error:', logError.message);
            }
        }
        
        const response = { 
            error: error.message,
            code: errorCode,
            timestamp: new Date().toISOString()
        };
        
        // Add debug info in development
        if (process.env.NODE_ENV === 'development') {
            response.stack = error.stack;
            response.fileInfo = req.file ? {
                originalname: req.file.originalname,
                mimetype: req.file.mimetype,
                size: req.file.size
            } : null;
        }
        
        res.status(statusCode).json(response);
    }
});

this.app.post('/api/debug-file', this.upload.single('template'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }
        
        const documentGenerator = new DocumentGenerator();
        await documentGenerator.validateDocxBuffer(req.file.buffer);
        
        // Try to parse the DOCX
        const PizZip = require('pizzip');
        const zip = new PizZip(req.file.buffer);
        const files = Object.keys(zip.files);
        
        // Try to read document.xml
        let documentContent = '';
        let hasVariables = false;
        
        if (zip.files['word/document.xml']) {
            documentContent = zip.files['word/document.xml'].asText();
            hasVariables = /\$\{[^}]+\}/.test(documentContent);
        }
        
        res.json({
            filename: req.file.originalname,
            mimetype: req.file.mimetype,
            size: req.file.size,
            isValidDocx: true,
            filesInArchive: files,
            documentLength: documentContent.length,
            hasTemplateVariables: hasVariables,
            sampleContent: documentContent.substring(0, 500) + (documentContent.length > 500 ? '...' : '')
        });
        
    } catch (error) {
        res.status(400).json({
            filename: req.file ? req.file.originalname : 'none',
            isValidDocx: false,
            error: error.message
        });
    }
});

// Bulk generation endpoint - multiple JSON datasets with one template
this.app.post('/api/generate-bulk', this.upload.single('template'), async (req, res) => {
    try {
        console.log('Bulk generation request received');
        
        if (!req.file) {
            return res.status(400).json({ 
                error: 'No template file uploaded' 
            });
        }

        if (!req.body.datasets) {
            return res.status(400).json({ 
                error: 'No datasets provided. Please include an array of JSON objects in the "datasets" field.' 
            });
        }

        // Parse datasets
        let datasets;
        try {
            datasets = typeof req.body.datasets === 'string' ? 
                JSON.parse(req.body.datasets) : req.body.datasets;
        } catch (parseError) {
            return res.status(400).json({ 
                error: 'Invalid datasets JSON: ' + parseError.message 
            });
        }

        if (!Array.isArray(datasets)) {
            return res.status(400).json({ 
                error: 'Datasets must be an array of JSON objects' 
            });
        }

        if (datasets.length === 0) {
            return res.status(400).json({ 
                error: 'At least one dataset is required' 
            });
        }

        if (datasets.length > 100) {
            return res.status(400).json({ 
                error: 'Maximum 100 datasets allowed per bulk request' 
            });
        }

        console.log(`Processing ${datasets.length} datasets with template:`, req.file.originalname);

        // Initialize document generator
        const documentGenerator = new DocumentGenerator();
        
        const options = {
            outputFormat: req.body.outputFormat || 'docx',
            batchSize: Math.min(parseInt(req.body.batchSize) || 10, 50),
            batchDelay: parseInt(req.body.batchDelay) || 100
        };

        // Generate documents
        const results = await documentGenerator.batchGenerate(
            req.file.buffer, 
            datasets, 
            options
        );

        // Prepare response
        const successful = results.filter(r => r.success);
        const failed = results.filter(r => !r.success);

        console.log(`Bulk generation complete: ${successful.length} successful, ${failed.length} failed`);

        // For bulk operations, return a summary and download links for successful generations
        const response = {
            summary: {
                total: datasets.length,
                successful: successful.length,
                failed: failed.length,
                totalSize: successful.reduce((sum, r) => sum + (r.size || 0), 0)
            },
            results: results.map(r => ({
                index: r.index,
                success: r.success,
                error: r.error,
                size: r.size
            })),
            // If only one document, return it directly
            singleDocument: successful.length === 1
        };

        if (successful.length === 1) {
            // Return single document directly
            const result = successful[0];
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
            const filename = `${req.file.originalname.replace('.docx', '')}_bulk_${timestamp}.docx`;

            res.set({
                'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'Content-Disposition': `attachment; filename="${filename}"`,
                'X-Bulk-Results': JSON.stringify(response.summary)
            });

            res.send(result.document);
        } else {
            // Return summary for multiple documents
            res.json(response);
        }

    } catch (error) {
        console.error('Bulk generation error:', error);
        res.status(500).json({ 
            error: 'Bulk generation failed: ' + error.message 
        });
    }
});

// Template validation endpoint - validate template without storing
this.app.post('/api/validate-template', this.upload.single('template'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ 
                error: 'No template file uploaded' 
            });
        }

        const templateValidator = new TemplateValidator();
        const validation = await templateValidator.validateTemplate(req.file.buffer);
        
        const response = {
            filename: req.file.originalname,
            fileSize: req.file.size,
            validation: validation,
            summary: {
                isValid: validation.valid,
                complexity: validation.statistics.complexityScore,
                features: {
                    variables: validation.statistics.totalPlaceholders,
                    conditions: validation.statistics.totalConditions,
                    loops: validation.statistics.totalLoops,
                    tables: validation.statistics.totalTables,
                    aggregations: validation.statistics.totalAggregations
                }
            }
        };

        res.json(response);

    } catch (error) {
        console.error('Template validation error:', error);
        res.status(500).json({ 
            error: 'Template validation failed: ' + error.message 
        });
    }
});

// Health check endpoint
this.app.get('/api/health', (req, res) => {
    res.json({
        status: 'healthy',
        timestamp: new Date().toISOString(),
        version: '1.0.0',
        features: {
            directGeneration: true,
            bulkGeneration: true,
            templateValidation: true,
            templateStorage: true
        },
        limits: {
            maxFileSize: '10MB',
            maxBulkDatasets: 100,
            supportedFormats: ['.docx']
        }
    });
});

// API documentation endpoint
this.app.get('/api/docs', (req, res) => {
    res.json({
        title: 'DOCX Template Engine API',
        version: '1.0.0',
        description: 'Advanced DOCX template processing with variables, conditions, loops, and tables',
        endpoints: {
            'POST /api/generate-direct': {
                description: 'Upload template + JSON data → Generate DOCX in one request',
                parameters: {
                    template: 'multipart/form-data file (.docx)',
                    data: 'JSON string or object with template data',
                    addPageNumbers: 'boolean (optional)',
                    outputFormat: 'string (optional, default: "docx")'
                },
                returns: 'Generated DOCX file'
            },
            'POST /api/generate-bulk': {
                description: 'Generate multiple documents from one template with different datasets',
                parameters: {
                    template: 'multipart/form-data file (.docx)',
                    datasets: 'JSON array of data objects',
                    batchSize: 'number (optional, max: 50)',
                    batchDelay: 'number in ms (optional)'
                },
                returns: 'Single DOCX file or bulk results summary'
            },
            'POST /api/validate-template': {
                description: 'Validate template syntax and features without storing',
                parameters: {
                    template: 'multipart/form-data file (.docx)'
                },
                returns: 'Validation results with detected features'
            },
            'GET /api/health': {
                description: 'API health check',
                returns: 'System status and capabilities'
            }
        },
        examples: {
            curl_direct: `curl -X POST http://localhost:3000/api/generate-direct \\
  -F "template=@template.docx" \\
  -F "data={\\"user\\":{\\"name\\":\\"John\\"},\\"amount\\":1000}" \\
  -o generated.docx`,
            curl_validate: `curl -X POST http://localhost:3000/api/validate-template \\
  -F "template=@template.docx"`
        }
    });
});

console.log('✅ Added API endpoints:');
console.log('   POST /api/generate-direct - Direct template + data → DOCX generation');
console.log('   POST /api/generate-bulk - Bulk document generation');
console.log('   POST /api/validate-template - Template validation');
console.log('   GET /api/health - Health check');
console.log('   GET /api/docs - API documentation');

        // Error handling
        this.app.use((error, req, res, next) => {
            if (error instanceof multer.MulterError) {
                return res.status(400).json({ error: error.message });
            }
            res.status(500).json({ error: 'Internal server error' });
        });
    }

    hashData(data) {
        const crypto = require('crypto');
        return crypto.createHash('md5').update(JSON.stringify(data)).digest('hex');
    }

    async start() {
        await this.connectDatabase();
        this.app.listen(this.port, () => {
            console.log(`DOCX Template Engine running on port ${this.port}`);
        });
    }
}

// Start server
const server = new DocxTemplateServer();
server.start().catch(console.error);

module.exports = DocxTemplateServer;