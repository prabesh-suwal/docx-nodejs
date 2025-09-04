const { v4: uuidv4 } = require('uuid');
const TemplateValidator = require('../validators/TemplateValidator');

class TemplateManager {
    constructor() {
        this.validator = new TemplateValidator();
    }

    async uploadTemplate(db, templateData) {
        const templateId = uuidv4();
        const timestamp = new Date();

        // Validate the template
        const validation = await this.validator.validateTemplate(templateData.buffer);
        
        const template = {
            _id: templateId,
            name: templateData.name,
            description: templateData.description,
            author: templateData.author,
            version: 1,
            createdAt: timestamp,
            updatedAt: timestamp,
            fileSize: templateData.buffer.length,
            validation: validation,
            metadata: {
                placeholders: validation.placeholders,
                conditions: validation.conditions,
                loops: validation.loops,
                tables: validation.tables,
                aggregations: validation.aggregations
            }
        };

        // Store template metadata
        await db.collection('templates').insertOne(template);

        // Store template file in GridFS or similar
        await db.collection('template_files').insertOne({
            templateId: templateId,
            buffer: templateData.buffer,
            createdAt: timestamp
        });

        return {
            id: templateId,
            ...template,
            buffer: undefined // Don't return buffer in response
        };
    }

    async updateTemplate(db, templateId, templateData) {
        const existingTemplate = await this.getTemplate(db, templateId);
        if (!existingTemplate) {
            throw new Error('Template not found');
        }

        // Validate new template
        const validation = await this.validator.validateTemplate(templateData.buffer);
        
        const updatedTemplate = {
            name: templateData.name || existingTemplate.name,
            description: templateData.description || existingTemplate.description,
            version: existingTemplate.version + 1,
            updatedAt: new Date(),
            fileSize: templateData.buffer.length,
            validation: validation,
            metadata: {
                placeholders: validation.placeholders,
                conditions: validation.conditions,
                loops: validation.loops,
                tables: validation.tables,
                aggregations: validation.aggregations
            }
        };

        // Update metadata
        await db.collection('templates').updateOne(
            { _id: templateId },
            { $set: updatedTemplate }
        );

        // Update file
        await db.collection('template_files').updateOne(
            { templateId: templateId },
            { 
                $set: {
                    buffer: templateData.buffer,
                    updatedAt: new Date()
                }
            },
            { upsert: true }
        );

        return this.getTemplate(db, templateId);
    }

    async listTemplates(db, options = {}) {
        const { limit = 50, skip = 0, sortBy = 'createdAt', sortOrder = -1 } = options;
        
        const templates = await db.collection('templates')
            .find({}, { projection: { buffer: 0 } })
            .sort({ [sortBy]: sortOrder })
            .skip(skip)
            .limit(limit)
            .toArray();

        const total = await db.collection('templates').countDocuments();

        return {
            templates,
            pagination: {
                total,
                limit,
                skip,
                hasMore: skip + limit < total
            }
        };
    }

    async getTemplate(db, templateId) {
        const template = await db.collection('templates').findOne({ _id: templateId });
        if (!template) return null;

        const file = await db.collection('template_files').findOne({ templateId });
        
        return {
            ...template,
            buffer: file ? file.buffer : null
        };
    }

    async deleteTemplate(db, templateId) {
        // Delete metadata
        const deleteResult = await db.collection('templates').deleteOne({ _id: templateId });
        
        if (deleteResult.deletedCount === 0) {
            throw new Error('Template not found');
        }

        // Delete file
        await db.collection('template_files').deleteOne({ templateId });

        // Delete related document logs
        await db.collection('document_logs').deleteMany({ templateId });

        return { success: true, message: 'Template deleted successfully' };
    }

    async searchTemplates(db, query, options = {}) {
        const { limit = 20, skip = 0 } = options;
        
        const searchFilter = {
            $or: [
                { name: { $regex: query, $options: 'i' } },
                { description: { $regex: query, $options: 'i' } },
                { author: { $regex: query, $options: 'i' } }
            ]
        };

        const templates = await db.collection('templates')
            .find(searchFilter, { projection: { buffer: 0 } })
            .sort({ createdAt: -1 })
            .skip(skip)
            .limit(limit)
            .toArray();

        return templates;
    }
}

module.exports = TemplateManager;
