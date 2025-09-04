const PizZip = require('pizzip');
const xml2js = require('xml2js');

class TemplateValidator {
    constructor() {
        this.parser = new xml2js.Parser();
    }

    async validateTemplate(templateBuffer) {
        try {
            const zip = new PizZip(templateBuffer);
            const documentXml = zip.files['word/document.xml'].asText();
            
            const validation = {
                valid: true,
                warnings: [],
                errors: [],
                placeholders: [],
                conditions: [],
                loops: [],
                tables: [],
                aggregations: [],
                formatting: [],
                statistics: {}
            };

            // Extract and validate placeholders
            await this.validatePlaceholders(documentXml, validation);
            
            // Validate conditions
            await this.validateConditions(documentXml, validation);
            
            // Validate loops
            await this.validateLoops(documentXml, validation);
            
            // Validate tables
            await this.validateTables(documentXml, validation);
            
            // Validate aggregations
            await this.validateAggregations(documentXml, validation);
            
            // Check for formatting issues
            await this.validateFormatting(documentXml, validation);
            
            // Generate statistics
            this.generateStatistics(validation);
            
            return validation;
        } catch (error) {
            return {
                valid: false,
                errors: [`Template validation failed: ${error.message}`],
                warnings: [],
                placeholders: [],
                conditions: [],
                loops: [],
                tables: [],
                aggregations: [],
                formatting: [],
                statistics: {}
            };
        }
    }

    async validatePlaceholders(xml, validation) {
        // Find all variable placeholders: ${...}
        const placeholderRegex = /\$\{([^}]+)\}/g;
        const matches = [...xml.matchAll(placeholderRegex)];
        
        const placeholderMap = new Map();
        
        for (const match of matches) {
            const fullMatch = match[0];
            const expression = match[1].trim();
            
            // Parse expression and formatters
            const parts = expression.split('|');
            const variable = parts[0].trim();
            const formatters = parts.slice(1).map(f => f.trim());
            
            const placeholder = {
                raw: fullMatch,
                variable: variable,
                formatters: formatters,
                valid: true,
                warnings: [],
                errors: []
            };
            
            // Validate variable syntax
            this.validateVariableSyntax(placeholder);
            
            // Validate formatters
            this.validateFormatters(placeholder);
            
            // Check for Unicode issues
            this.checkUnicodeIssues(placeholder);
            
            placeholderMap.set(fullMatch, placeholder);
        }
        
        validation.placeholders = Array.from(placeholderMap.values());
        
        // Collect warnings and errors
        validation.placeholders.forEach(p => {
            validation.warnings.push(...p.warnings);
            validation.errors.push(...p.errors);
            if (p.errors.length > 0) validation.valid = false;
        });
    }

    validateVariableSyntax(placeholder) {
        const variable = placeholder.variable;
        
        // Check for valid variable name
        if (!/^[a-zA-Z_][a-zA-Z0-9_.$\[\]]*$/.test(variable)) {
            placeholder.errors.push(`Invalid variable syntax: ${variable}`);
            placeholder.valid = false;
        }
        
        // Check for array access syntax
        const arrayMatch = variable.match(/\[(\d+)\]/);
        if (arrayMatch) {
            const index = parseInt(arrayMatch[1]);
            if (index < 0) {
                placeholder.warnings.push(`Negative array index: ${variable}`);
            }
        }
        
        // Check for deeply nested properties
        const depth = variable.split('.').length;
        if (depth > 10) {
            placeholder.warnings.push(`Deeply nested property (${depth} levels): ${variable}`);
        }
    }

    validateFormatters(placeholder) {
        const validFormatters = [
            'upper', 'lower', 'capitalize', 'trim',
            'currency', 'number', 'percent', 'round',
            'date', 'dateTime', 'fromNow',
            'join', 'length', 'sum', 'count', 'avg',
            'truncate', 'default', 'escape',
            'bold', 'italic', 'underline', 'size', 'color'
        ];
        
        for (const formatter of placeholder.formatters) {
            const [name] = formatter.split(':');
            
            if (!validFormatters.includes(name)) {
                placeholder.warnings.push(`Unknown formatter: ${name}`);
            }
            
            // Validate formatter parameters
            this.validateFormatterParameters(formatter, placeholder);
        }
    }

    validateFormatterParameters(formatter, placeholder) {
        const [name, ...params] = formatter.split(':');
        
        switch (name) {
            case 'date':
            case 'dateTime':
                if (params.length > 0 && !this.isValidDateFormat(params[0])) {
                    placeholder.warnings.push(`Invalid date format: ${params[0]}`);
                }
                break;
            case 'currency':
                if (params.length > 0 && !this.isValidCurrencyCode(params[0])) {
                    placeholder.warnings.push(`Invalid currency code: ${params[0]}`);
                }
                break;
            case 'round':
            case 'number':
                if (params.length > 0 && isNaN(parseInt(params[0]))) {
                    placeholder.warnings.push(`Invalid number parameter: ${params[0]}`);
                }
                break;
            case 'size':
                if (params.length > 0) {
                    const size = parseInt(params[0]);
                    if (isNaN(size) || size < 1 || size > 72) {
                        placeholder.warnings.push(`Invalid font size: ${params[0]} (must be 1-72)`);
                    }
                }
                break;
        }
    }

    checkUnicodeIssues(placeholder) {
        const text = placeholder.raw;
        
        // Check for smart quotes and other problematic characters
        const problematicChars = /[""''–—]/;
        if (problematicChars.test(text)) {
            placeholder.warnings.push('Contains smart quotes or special characters that may cause issues');
        }
        
        // Check for invisible characters
        const invisibleChars = /[\u200B-\u200D\uFEFF]/;
        if (invisibleChars.test(text)) {
            placeholder.warnings.push('Contains invisible Unicode characters');
        }
    }

    async validateConditions(xml, validation) {
        const conditionRegex = /\$\{#if\s+([^}]+)\}([\s\S]*?)(?:\$\{#else\}([\s\S]*?))?\$\{\/if\}/g;
        const matches = [...xml.matchAll(conditionRegex)];
        
        for (const match of matches) {
            const condition = {
                raw: match[0],
                expression: match[1].trim(),
                ifContent: match[2],
                elseContent: match[3] || '',
                valid: true,
                warnings: [],
                errors: []
            };
            
            // Validate condition syntax
            this.validateConditionSyntax(condition);
            
            validation.conditions.push(condition);
        }
    }

    validateConditionSyntax(condition) {
        const expr = condition.expression;
        
        // Check for valid operators
        const validOperators = ['==', '!=', '>', '<', '>=', '<=', '&&', '||', 'and', 'or', 'not'];
        const hasOperator = validOperators.some(op => expr.includes(op));
        
        if (!hasOperator && !expr.includes(' ')) {
            condition.warnings.push('Simple boolean condition, consider being more explicit');
        }
        
        // Check for balanced parentheses
        let parenCount = 0;
        for (const char of expr) {
            if (char === '(') parenCount++;
            if (char === ')') parenCount--;
        }
        
        if (parenCount !== 0) {
            condition.errors.push('Unbalanced parentheses in condition');
            condition.valid = false;
        }
    }

    async validateLoops(xml, validation) {
        const loopRegex = /\$\{#each\s+([^}]+)\}([\s\S]*?)\$\{\/each\}/g;
        const matches = [...xml.matchAll(loopRegex)];
        
        for (const match of matches) {
            const loop = {
                raw: match[0],
                array: match[1].trim(),
                content: match[2],
                nested: false,
                valid: true,
                warnings: [],
                errors: []
            };
            
            // Check for nested loops
            if (loopRegex.test(loop.content)) {
                loop.nested = true;
                loop.warnings.push('Contains nested loops - ensure data structure supports this');
            }
            
            // Validate array reference
            this.validateArrayReference(loop);
            
            validation.loops.push(loop);
        }
    }

    validateArrayReference(loop) {
        const arrayRef = loop.array;
        
        // Check for valid array syntax
        if (!/^[a-zA-Z_][a-zA-Z0-9_.$\[\]]*$/.test(arrayRef)) {
            loop.errors.push(`Invalid array reference: ${arrayRef}`);
            loop.valid = false;
        }
        
        // Check for 'this' usage in loop content
        if (loop.content.includes('${this.') && !loop.content.includes('this')) {
            loop.warnings.push('Uses "this" context - ensure loop data has required properties');
        }
    }

    async validateTables(xml, validation) {
        // Find table structures with template variables
        const tableRegex = /<w:tbl[^>]*>([\s\S]*?)<\/w:tbl>/g;
        const tableMatches = [...xml.matchAll(tableRegex)];
        
        for (const tableMatch of tableMatches) {
            const tableContent = tableMatch[1];
            
            // Check if table contains template variables
            if (/\$\{/.test(tableContent)) {
                const table = {
                    raw: tableMatch[0],
                    content: tableContent,
                    hasVariables: true,
                    hasLoops: false,
                    expandable: false,
                    warnings: [],
                    errors: []
                };
                
                // Check for loop-expandable tables
                if (/\$\{#each/.test(tableContent)) {
                    table.hasLoops = true;
                    table.expandable = true;
                }
                
                validation.tables.push(table);
            }
        }
    }

    async validateAggregations(xml, validation) {
        const aggregationRegex = /\$\{([^}]*)\|(sum|count|avg|max|min)([^}]*)\}/g;
        const matches = [...xml.matchAll(aggregationRegex)];
        
        for (const match of matches) {
            const aggregation = {
                raw: match[0],
                variable: match[1].trim(),
                operation: match[2],
                parameters: match[3] ? match[3].replace(':', '').trim() : '',
                valid: true,
                warnings: [],
                errors: []
            };
            
            // Validate aggregation syntax
            this.validateAggregationSyntax(aggregation);
            
            validation.aggregations.push(aggregation);
        }
    }

    validateAggregationSyntax(aggregation) {
        const validOps = ['sum', 'count', 'avg', 'max', 'min'];
        
        if (!validOps.includes(aggregation.operation)) {
            aggregation.errors.push(`Invalid aggregation operation: ${aggregation.operation}`);
            aggregation.valid = false;
        }
        
        if (['sum', 'avg', 'max', 'min'].includes(aggregation.operation) && !aggregation.parameters) {
            aggregation.warnings.push(`${aggregation.operation} operation may need a field parameter`);
        }
    }

    async validateFormatting(xml, validation) {
        const formattingRegex = /\$\{[^}]*\|(bold|italic|underline|size|color)[^}]*\}/g;
        const matches = [...xml.matchAll(formattingRegex)];
        
        for (const match of matches) {
            const formatting = {
                raw: match[0],
                type: match[1],
                valid: true,
                warnings: [],
                errors: []
            };
            
            validation.formatting.push(formatting);
        }
    }

    generateStatistics(validation) {
        validation.statistics = {
            totalPlaceholders: validation.placeholders.length,
            uniqueVariables: new Set(validation.placeholders.map(p => p.variable)).size,
            totalConditions: validation.conditions.length,
            totalLoops: validation.loops.length,
            nestedLoops: validation.loops.filter(l => l.nested).length,
            totalTables: validation.tables.length,
            expandableTables: validation.tables.filter(t => t.expandable).length,
            totalAggregations: validation.aggregations.length,
            formattingDirectives: validation.formatting.length,
            totalWarnings: validation.warnings.length,
            totalErrors: validation.errors.length,
            complexityScore: this.calculateComplexityScore(validation)
        };
    }

    calculateComplexityScore(validation) {
        let score = 0;
        score += validation.placeholders.length * 1;
        score += validation.conditions.length * 3;
        score += validation.loops.length * 5;
        score += validation.loops.filter(l => l.nested).length * 10;
        score += validation.aggregations.length * 4;
        score += validation.formatting.length * 2;
        
        return score;
    }

    isValidDateFormat(format) {
        // Basic validation for moment.js date formats
        const validPatterns = /^[YMDHmsaA\/\-\s:.,]*$/;
        return validPatterns.test(format);
    }

    isValidCurrencyCode(code) {
        const validCurrencies = ['USD', 'EUR', 'GBP', 'JPY', 'CAD', 'AUD', 'CHF', 'CNY', 'INR'];
        return validCurrencies.includes(code.toUpperCase());
    }
}

module.exports = TemplateValidator;