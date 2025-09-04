const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const FormatHelper = require('./FormatHelper');
const ExpressionEvaluator = require('./ExpressionEvaluator');

class TemplateEngine {
    constructor() {
        this.formatHelper = new FormatHelper();
        this.expressionEvaluator = new ExpressionEvaluator();
        
        // Register custom modules
        this.modules = [];
        this.setupDefaultModules();
    }

    setupDefaultModules() {
        // Custom Variable Module
        const variableModule = {
            name: 'VariableModule',
            prefix: '\\$\\{',
            suffix: '\\}',
            
            parse: (tag) => {
                const parts = tag.value.split('|');
                const expression = parts[0].trim();
                const formatters = parts.slice(1).map(f => f.trim());
                
                return {
                    value: expression,
                    formatters: formatters,
                    module: 'VariableModule'
                };
            },

            render: (part, { data }) => {
                try {
                    let value = this.expressionEvaluator.evaluate(part.value, data);
                    
                    // Apply formatters
                    if (part.formatters && part.formatters.length > 0) {
                        value = this.formatHelper.applyFormatters(value, part.formatters);
                    }
                    
                    return value || '';
                } catch (error) {
                    console.warn(`Variable evaluation error for "${part.value}":`, error.message);
                    return `[ERROR: ${part.value}]`;
                }
            }
        };

        // Custom Condition Module
        const conditionModule = {
            name: 'ConditionModule',
            prefix: '\\$\\{#if\\s+',
            suffix: '\\}',
            
            parse: (tag) => {
                return {
                    condition: tag.value.replace(/^#if\s+/, '').trim(),
                    module: 'ConditionModule'
                };
            },

            render: (part, { data }) => {
                try {
                    const result = this.expressionEvaluator.evaluateCondition(part.condition, data);
                    return result;
                } catch (error) {
                    console.warn(`Condition evaluation error for "${part.condition}":`, error.message);
                    return false;
                }
            }
        };

        // Custom Loop Module
        const loopModule = {
            name: 'LoopModule',
            prefix: '\\$\\{#each\\s+',
            suffix: '\\}',
            
            parse: (tag) => {
                return {
                    array: tag.value.replace(/^#each\s+/, '').trim(),
                    module: 'LoopModule'
                };
            },

            render: (part, { data }) => {
                try {
                    const arrayData = this.expressionEvaluator.evaluate(part.array, data);
                    return Array.isArray(arrayData) ? arrayData : [];
                } catch (error) {
                    console.warn(`Loop evaluation error for "${part.array}":`, error.message);
                    return [];
                }
            }
        };

        this.modules.push(variableModule, conditionModule, loopModule);
    }

    async processTemplate(templateBuffer, data) {
    try {
        console.log('🚀 Starting template processing...');
        return await this.processAdvancedTemplate(templateBuffer, data);
    } catch (error) {
        throw new Error(`Template processing failed: ${error.message}`);
    }
}

    preprocessData(data) {
        // Deep clone the data to avoid mutations
        const processedData = JSON.parse(JSON.stringify(data));
        
        // Add helper functions to data context
        processedData._helpers = {
            sum: (array, field) => this.aggregateHelper.sum(array, field),
            count: (array) => this.aggregateHelper.count(array),
            avg: (array, field) => this.aggregateHelper.average(array, field),
            max: (array, field) => this.aggregateHelper.max(array, field),
            min: (array, field) => this.aggregateHelper.min(array, field)
        };
        
        return processedData;
    }




    // Custom template processing for complex syntax
    async processAdvancedTemplate(templateBuffer, data) {
    try {
        const zip = new PizZip(templateBuffer);
        
        // Extract document.xml
        const documentXml = zip.files['word/document.xml'].asText();
        
        console.log('🔄 Starting enhanced template processing pipeline...');
        
        // ENHANCED PROCESSING ORDER:
        // 1. Process loops first (including advanced table loops)
        // 2. Process advanced tables with loop marker removal
        // 3. Process conditions
        // 4. Clean up remaining variables
        
        console.log('📝 Step 1: Processing loops...');
        let processedXml = this.processLoops(documentXml, data);
        
        console.log('📝 Step 2: Processing advanced tables...');
        processedXml = this.processAdvancedTable(processedXml, data);
        
        console.log('📝 Step 3: Processing tables and removing empty control rows...');
        processedXml = this.processTables(processedXml, data);
        
        console.log('📝 Step 4: Processing conditions...');
        processedXml = this.processConditions(processedXml, data);
        
        console.log('📝 Step 5: Processing remaining variables...');
        processedXml = this.processRemainingVariables(processedXml, data);
        
        console.log('✅ Enhanced template processing complete');
        
        // Update the zip with processed content
        zip.file('word/document.xml', processedXml);
        
        return zip.generate({ type: 'nodebuffer' });
    } catch (error) {
        throw new Error(`Advanced template processing failed: ${error.message}`);
    }
}

processRemainingVariables(xml, data) {
    const variableRegex = /\$\{([^}]+)\}/g;
    
    return xml.replace(variableRegex, (match, expression) => {
        try {
            // Skip template control structures (they should already be processed)
            if (expression.includes('#each') || expression.includes('#if') || expression.includes('/each') || expression.includes('/if')) {
                console.log(`Skipping control structure: ${expression}`);
                return match; // Leave as-is
            }
            
            const parts = expression.split('|');
            const varPath = parts[0].trim();
            const formatters = parts.slice(1).map(f => f.trim());
            
            // Skip 'this.' variables (should be processed in loops)
            if (varPath.startsWith('this.')) {
                console.log(`Warning: Found unprocessed loop variable: ${varPath}`);
                return ''; // Remove unprocessed loop variables
            }
            
            console.log(`Processing remaining variable: ${varPath}`);
            let value = this.expressionEvaluator.evaluate(varPath, data);
            
            // Apply formatters
            if (formatters.length > 0) {
                value = this.formatHelper.applyFormatters(value, formatters);
            }
            
            // Handle formatting objects
            if (value && typeof value === 'object' && value.formatting) {
                return this.escapeXml(String(value.value || ''));
            }
            
            // Handle XML entities
            return this.escapeXml(String(value || ''));
        } catch (error) {
            console.warn(`Remaining variable processing error for "${expression}":`, error.message);
            return `[ERROR: ${expression}]`;
        }
    });
}

    processVariables(xml, data) {
    // If we're not in a loop context, use the regular processing
    if (!data.this) {
        return this.processRegularVariables(xml, data);
    }
    
    // We're in a loop context, use enhanced processing
    return this.processVariablesInLoop(xml, data);
}

processRegularVariables(xml, data) {
    const variableRegex = /\$\{([^}]+)\}/g;
    
    return xml.replace(variableRegex, (match, expression) => {
        try {
            const parts = expression.split('|');
            const varPath = parts[0].trim();
            const formatters = parts.slice(1).map(f => f.trim());
            
            let value = this.expressionEvaluator.evaluate(varPath, data);
            
            // Apply formatters
            if (formatters.length > 0) {
                value = this.formatHelper.applyFormatters(value, formatters);
            }
            
            // Handle formatting objects
            if (value && typeof value === 'object' && value.formatting) {
                return this.escapeXml(String(value.value || ''));
            }
            
            // Handle XML entities
            return this.escapeXml(String(value || ''));
        } catch (error) {
            console.warn(`Variable processing error for "${expression}":`, error.message);
            return `[ERROR: ${expression}]`;
        }
    });
}

    processConditions(xml, data) {
    // Match conditions: ${#if condition}...${#else}...${/if}
    const conditionRegex = /\$\{#if\s+([^}]+)\}([\s\S]*?)(?:\$\{#else\}([\s\S]*?))?\$\{\/if\}/g;
    
    return xml.replace(conditionRegex, (match, condition, ifContent, elseContent = '') => {
        try {
            console.log(`Processing condition: ${condition}`);
            
            // Enhanced condition evaluation for loop context
            let result;
            
            // Handle 'this' context in conditions
            const processedCondition = condition.replace(/\bthis\.(\w+)/g, (match, prop) => {
                if (data.this && data.this[prop] !== undefined) {
                    const value = data.this[prop];
                    return typeof value === 'string' ? `"${value}"` : String(value);
                }
                return 'null';
            });
            
            result = this.expressionEvaluator.evaluateCondition(processedCondition, data);
            
            console.log(`Condition "${condition}" evaluated to: ${result}`);
            return result ? ifContent : elseContent;
            
        } catch (error) {
            console.error(`Condition processing error for "${condition}":`, error);
            return ifContent; // Default to showing content on error
        }
    });
}


    processLoops(xml, data) {
    const loopRegex = /\$\{#each\s+([^}]+)\}([\s\S]*?)\$\{\/each\}/g;
    
    return xml.replace(loopRegex, (match, arrayPath, content) => {
        try {
            console.log(`Processing loop for array: ${arrayPath}`);
            const arrayData = this.expressionEvaluator.evaluate(arrayPath, data);
            
            if (!Array.isArray(arrayData)) {
                console.warn(`Loop data is not an array for path: ${arrayPath}`, arrayData);
                return '';
            }
            
            console.log(`Found ${arrayData.length} items in loop: ${arrayPath}`);
            
            return arrayData.map((item, index) => {
                // Create enhanced context data with proper 'this' reference
                const contextData = {
                    ...data,
                    this: item,
                    index: index,
                    first: index === 0,
                    last: index === arrayData.length - 1,
                    count: arrayData.length
                };
                
                console.log(`Processing loop item ${index}:`, Object.keys(item));
                
                // Process ALL template elements within the loop context
                let processedContent = content;
                
                // 1. Process variables first (with loop context)
                processedContent = this.processVariablesInLoop(processedContent, contextData);
                
                // 2. Process conditions (with loop context)  
                processedContent = this.processConditionsInLoop(processedContent, contextData);
                
                // 3. Handle nested loops recursively
                processedContent = this.processLoops(processedContent, contextData);
                
                return processedContent;
            }).join('');
            
        } catch (error) {
            console.error(`Loop processing error for "${arrayPath}":`, error);
            return `[ERROR: Loop ${arrayPath} - ${error.message}]`;
        }
    });
}

processConditionsInLoop(xml, data) {
    const conditionRegex = /\$\{#if\s+([^}]+)\}([\s\S]*?)(?:\$\{#else\}([\s\S]*?))?\$\{\/if\}/g;
    
    return xml.replace(conditionRegex, (match, condition, ifContent, elseContent = '') => {
        try {
            console.log(`Processing condition in loop: ${condition}`);
            
            // Enhanced condition evaluation for loop context
            let result = this.evaluateConditionInLoop(condition, data);
            
            console.log(`Loop condition "${condition}" evaluated to: ${result}`);
            return result ? ifContent : elseContent;
            
        } catch (error) {
            console.error(`Loop condition processing error for "${condition}":`, error);
            return ifContent; // Default to showing content on error
        }
    });
}

evaluateConditionInLoop(condition, data) {
    // Handle 'this.property' references in conditions
    let processedCondition = condition;
    
    // Replace this.property with actual values
    processedCondition = processedCondition.replace(/\bthis\.([a-zA-Z_][a-zA-Z0-9_]*)\b/g, (match, property) => {
        if (data.this && data.this[property] !== undefined) {
            const value = data.this[property];
            const serialized = typeof value === 'string' ? `"${value}"` : String(value);
            console.log(`Replacing ${match} with ${serialized} in condition`);
            return serialized;
        }
        console.log(`Property ${property} not found in this context for condition`);
        return 'null';
    });
    
    // Simple evaluation for common comparison operators
    try {
        // Handle >= <= == != comparisons
        const comparisonMatch = processedCondition.match(/^(.+?)\s*(>=|<=|==|!=|>|<)\s*(.+?)$/);
        if (comparisonMatch) {
            const [, left, operator, right] = comparisonMatch;
            const leftVal = this.parseValue(left.trim());
            const rightVal = this.parseValue(right.trim());
            
            console.log(`Comparing: ${leftVal} ${operator} ${rightVal}`);
            
            switch (operator) {
                case '>=': return leftVal >= rightVal;
                case '<=': return leftVal <= rightVal;
                case '>': return leftVal > rightVal;
                case '<': return leftVal < rightVal;
                case '==': return leftVal == rightVal;
                case '!=': return leftVal != rightVal;
            }
        }
        
        // Fallback to simple truthiness
        const value = this.parseValue(processedCondition);
        return Boolean(value);
        
    } catch (error) {
        console.error(`Condition evaluation error: ${error.message}`);
        return false;
    }
}


parseValue(str) {
    const trimmed = str.trim();
    
    // Number
    if (/^-?\d+\.?\d*$/.test(trimmed)) {
        return parseFloat(trimmed);
    }
    
    // String (quoted)
    if ((trimmed.startsWith('"') && trimmed.endsWith('"')) || 
        (trimmed.startsWith("'") && trimmed.endsWith("'"))) {
        return trimmed.slice(1, -1);
    }
    
    // Boolean
    if (trimmed === 'true') return true;
    if (trimmed === 'false') return false;
    if (trimmed === 'null') return null;
    
    // Default to the string value
    return trimmed;
}

    
processVariablesInLoop(xml, data) {
    // Enhanced variable regex that handles 'this' context better
    const variableRegex = /\$\{([^}]+)\}/g;
    
    return xml.replace(variableRegex, (match, expression) => {
        try {
            const parts = expression.split('|');
            const varPath = parts[0].trim();
            const formatters = parts.slice(1).map(f => f.trim());
            
            console.log(`Processing variable in loop: ${varPath}`);
            
            let value;
            
            // Special handling for 'this' context in loops
            if (varPath.startsWith('this.')) {
                const propertyPath = varPath.substring(5); // Remove 'this.'
                value = this.getNestedProperty(data.this, propertyPath);
                console.log(`Loop variable ${varPath} = ${value}`);
            } else if (varPath === 'this') {
                value = data.this;
            } else if (varPath === 'index') {
                value = data.index;
            } else if (varPath === 'first') {
                value = data.first;
            } else if (varPath === 'last') {
                value = data.last;
            } else if (varPath === 'count') {
                value = data.count;
            } else {
                // Regular variable evaluation
                value = this.expressionEvaluator.evaluate(varPath, data);
            }
            
            // Apply formatters if present
            if (formatters.length > 0) {
                console.log(`Applying formatters: ${formatters.join(', ')}`);
                value = this.formatHelper.applyFormatters(value, formatters);
            }
            
            // Handle formatting objects with styling
            if (value && typeof value === 'object' && value.formatting) {
                // For now, just return the value (DOCX formatting would be applied differently)
                return this.escapeXml(String(value.value || ''));
            }
            
            return this.escapeXml(String(value !== undefined && value !== null ? value : ''));
            
        } catch (error) {
            console.error(`Variable processing error in loop for "${expression}":`, error);
            return `[ERROR: ${expression}]`;
        }
    });
}

getNestedProperty(obj, path) {
    if (!obj || !path) return null;
    
    // Handle array access like "topDeals[0].amount"
    const normalizedPath = path.replace(/\[(\d+)\]/g, '.$1');
    
    const keys = normalizedPath.split('.');
    let current = obj;
    
    for (const key of keys) {
        if (current === null || current === undefined) {
            return null;
        }
        
        // Handle array indices
        if (/^\d+$/.test(key)) {
            const index = parseInt(key);
            if (Array.isArray(current) && index < current.length) {
                current = current[index];
            } else {
                return null;
            }
        } else {
            current = current[key];
        }
    }
    
    return current;
}

    processTables(xml, data) {
    console.log('📊 Processing tables with enhanced loop handling...');
    
    // First, let's handle the standard table processing for loops within tables
    let processedXml = this.processTableLoops(xml, data);
    
    // Then clean up any remaining empty control rows
    processedXml = this.removeEmptyControlRows(processedXml);
    
    return processedXml;
    }

    processTableLoops(xml, data) {
    // Find table rows with template variables
    const tableRowRegex = /<w:tr[^>]*>([\s\S]*?)<\/w:tr>/g;
    
    return xml.replace(tableRowRegex, (match, rowContent) => {
        // Check if this row contains loop start markers
        const loopStartMatch = rowContent.match(/\$\{#each\s+([^}]+)\}/);
        if (loopStartMatch) {
            const arrayPath = loopStartMatch[1];
            console.log(`Found table loop start for: ${arrayPath}`);
            
            try {
                const arrayData = this.expressionEvaluator.evaluate(arrayPath, data);
                
                if (!Array.isArray(arrayData)) {
                    console.warn(`Table loop data is not an array: ${arrayPath}`);
                    return ''; // Remove the row
                }
                
                // This row starts a loop - mark it for processing but don't generate content yet
                return match; // Keep the row marker for now
                
            } catch (error) {
                console.warn(`Table loop processing error for "${arrayPath}":`, error.message);
                return ''; // Remove problematic row
            }
        }
        
        // Check if this row contains loop end markers
        if (rowContent.includes('${/each}')) {
            console.log('Found table loop end marker');
            return match; // Keep the row marker for now
        }
        
        return match; // Regular row, keep as-is
    });
}

// Alternative approach: Pre-process table markup before XML processing
preprocessTableMarkup(content) {
    console.log('🔧 Preprocessing table markup to handle loop markers...');
    
    // Convert table-style markup to a format that's easier to process
    // This handles the case where users write tables with | separators
    
    const tablePattern = /(\|[^|\n]+\|[\s\S]*?)(\|[\s\S]*?\|)/g;
    
    return content.replace(tablePattern, (match, ...args) => {
        // Check if this table block contains loop markers
        if (match.includes('${#each}') || match.includes('${/each}')) {
            console.log('Found table with loop markers, preprocessing...');
            
            // Remove loop marker rows that are just empty cells
            let processed = match;
            
            // Remove rows that are just loop markers with empty cells
            processed = processed.replace(/\|\s*\$\{#each[^}]+\}\s*\|[\s\n]*/g, '${#each $1}\n');
            processed = processed.replace(/\|\s*\$\{\/each\}\s*\|[\s\n]*/g, '${/each}\n');
            
            return processed;
        }
        
        return match;
    });
}

processTablesWithControlRowRemoval(xml, data) {
    console.log('Processing tables and removing control rows...');
    
    // First, process loops normally
    let processedXml = this.processLoops(xml, data);
    
    // Then remove empty table rows that were created by control markers
    processedXml = this.removeTableControlRows(processedXml);
    
    return processedXml;
}

removeTableControlRows(xml) {
    console.log('Removing empty table control rows...');
    
    // Pattern to match Word table rows
    const tableRowRegex = /<w:tr[^>]*>([\s\S]*?)<\/w:tr>/g;
    
    return xml.replace(tableRowRegex, (match, rowContent) => {
        // Check if this row contains only empty table cells
        // This happens when ${#each} or ${/each} markers are processed and removed
        
        // Extract all text content from table cells
        const cellTextMatches = rowContent.match(/<w:t[^>]*>([^<]*)<\/w:t>/g);
        
        if (cellTextMatches) {
            // Get the actual text content from all cells
            const allText = cellTextMatches
                .map(match => match.replace(/<[^>]*>/g, '').trim())
                .join('');
            
            // If all cells are empty, remove this row
            if (allText.length === 0) {
                console.log('Removing empty table row');
                return '';
            }
        }
        
        // Also check for rows that contain only whitespace
        const textContent = rowContent.replace(/<[^>]*>/g, '').trim();
        if (textContent.length === 0) {
            console.log('Removing whitespace-only table row');
            return '';
        }
        
        return match; // Keep rows with actual content
    });
}

// Alternative approach: Process table loops with row cleanup
processTableLoopsClean(xml, data) {
    // Find table loop patterns like:
    // |${#each array}|
    // |${this.prop}|${this.prop2}|
    // |${/each}|
    
    const tableLoopPattern = /(\|[^|\n]*\$\{#each\s+([^}]+)\}[^|\n]*\|[\s\n]*)([\s\S]*?)(\|[^|\n]*\$\{\/each\}[^|\n]*\|)/g;
    
    return xml.replace(tableLoopPattern, (match, startRow, arrayPath, content, endRow) => {
        try {
            console.log(`Processing table loop: ${arrayPath}`);
            
            const arrayData = this.expressionEvaluator.evaluate(arrayPath, data);
            
            if (!Array.isArray(arrayData)) {
                console.warn(`Not an array: ${arrayPath}`);
                return '';
            }
            
            // Process each item in the array
            const generatedRows = arrayData.map((item, index) => {
                const contextData = {
                    ...data,
                    this: item,
                    index: index
                };
                
                // Process the content between start and end markers
                let processedContent = this.processVariablesInLoop(content, contextData);
                return processedContent;
            }).join('');
            
            return generatedRows; // Return only the generated content, no control rows
            
        } catch (error) {
            console.error(`Table loop error: ${error.message}`);
            return '';
        }
    });
}

removeEmptyControlRows(xml) {
    console.log('🧹 Removing empty control rows from tables...');
    
    // Pattern to match table rows that contain only whitespace and processed loop markers
    const emptyControlRowPattern = /<w:tr[^>]*>\s*<w:tc[^>]*>\s*<w:p[^>]*>\s*<w:t[^>]*>\s*<\/w:t>\s*<\/w:p>\s*<\/w:tc>\s*<\/w:tr>/g;
    
    // Remove completely empty table rows
    let cleaned = xml.replace(emptyControlRowPattern, '');
    
    // Pattern for rows with only empty cells (multiple cells)
    const emptyMultiCellRowPattern = /<w:tr[^>]*>(\s*<w:tc[^>]*>\s*<w:p[^>]*>\s*<w:t[^>]*>\s*<\/w:t>\s*<\/w:p>\s*<\/w:tc>\s*)+<\/w:tr>/g;
    cleaned = cleaned.replace(emptyMultiCellRowPattern, '');
    
    // Pattern for rows that contain only whitespace in table cells
    const whitespaceOnlyRowPattern = /<w:tr[^>]*>([\s\S]*?)<\/w:tr>/g;
    cleaned = cleaned.replace(whitespaceOnlyRowPattern, (match, content) => {
        // Check if the row content contains only whitespace and empty table cell structures
        const hasActualContent = /<w:t[^>]*>([^<]+)<\/w:t>/.test(content);
        const actualText = content.match(/<w:t[^>]*>([^<]*)<\/w:t>/g);
        
        if (actualText) {
            const hasNonEmptyText = actualText.some(text => {
                const innerText = text.replace(/<[^>]*>/g, '').trim();
                return innerText.length > 0;
            });
            
            if (!hasNonEmptyText) {
                console.log('Removing empty table row');
                return ''; // Remove empty row
            }
        }
        
        return match; // Keep row with actual content
    });
    
    return cleaned;
}


processAdvancedTable(xml, data) {
    console.log('🔧 Processing advanced table with loop markers...');
    
    // Pattern to match table structures with loop markers
    const tableLoopPattern = /<w:tr[^>]*>[\s\S]*?\$\{#each\s+([^}]+)\}[\s\S]*?<\/w:tr>([\s\S]*?)<w:tr[^>]*>[\s\S]*?\$\{\/each\}[\s\S]*?<\/w:tr>/g;
    
    return xml.replace(tableLoopPattern, (match, arrayPath, contentBetween) => {
        try {
            console.log(`Processing advanced table loop for: ${arrayPath}`);
            const arrayData = this.expressionEvaluator.evaluate(arrayPath, data);
            
            if (!Array.isArray(arrayData)) {
                console.warn(`Advanced table loop data is not an array: ${arrayPath}`);
                return '';
            }
            
            console.log(`Generating ${arrayData.length} table rows for ${arrayPath}`);
            
            // Generate table rows for each data item
            return arrayData.map((item, index) => {
                const contextData = {
                    ...data,
                    this: item,
                    index: index,
                    first: index === 0,
                    last: index === arrayData.length - 1
                };
                
                // Process the content between loop markers with context
                let processedRow = this.processVariablesInLoop(contentBetween, contextData);
                processedRow = this.processConditionsInLoop(processedRow, contextData);
                
                return processedRow;
            }).join('');
            
        } catch (error) {
            console.error(`Advanced table processing error for "${arrayPath}":`, error);
            return `[ERROR: Table loop ${arrayPath}]`;
        }
    });
}

    escapeXml(text) {
        return text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    // Aggregate helper for sum, count, avg operations
    get aggregateHelper() {
        return {
            sum: (array, field) => {
                if (!Array.isArray(array)) return 0;
                return array.reduce((sum, item) => {
                    const value = field ? this.expressionEvaluator.evaluate(field, { this: item }) : item;
                    return sum + (Number(value) || 0);
                }, 0);
            },
            
            count: (array) => {
                return Array.isArray(array) ? array.length : 0;
            },
            
            average: (array, field) => {
                if (!Array.isArray(array) || array.length === 0) return 0;
                const sum = this.aggregateHelper.sum(array, field);
                return sum / array.length;
            },
            
            max: (array, field) => {
                if (!Array.isArray(array) || array.length === 0) return null;
                return Math.max(...array.map(item => {
                    const value = field ? this.expressionEvaluator.evaluate(field, { this: item }) : item;
                    return Number(value) || 0;
                }));
            },
            
            min: (array, field) => {
                if (!Array.isArray(array) || array.length === 0) return null;
                return Math.min(...array.map(item => {
                    const value = field ? this.expressionEvaluator.evaluate(field, { this: item }) : item;
                    return Number(value) || 0;
                }));
            }
        };
    }
}

module.exports = TemplateEngine;