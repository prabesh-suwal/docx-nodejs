class ExpressionEvaluator {
    constructor() {
        this.operators = {
            '==': (a, b) => a == b,
            '===': (a, b) => a === b,
            '!=': (a, b) => a != b,
            '!==': (a, b) => a !== b,
            '>': (a, b) => Number(a) > Number(b),
            '<': (a, b) => Number(a) < Number(b),
            '>=': (a, b) => Number(a) >= Number(b),
            '<=': (a, b) => Number(a) <= Number(b),
            '&&': (a, b) => a && b,
            '||': (a, b) => a || b,
            '+': (a, b) => Number(a) + Number(b),
            '-': (a, b) => Number(a) - Number(b),
            '*': (a, b) => Number(a) * Number(b),
            '/': (a, b) => Number(a) / Number(b),
            '%': (a, b) => Number(a) % Number(b)
        };
    }

    evaluate(expression, data) {
    if (!expression) return null;
    
    console.log(`Evaluating expression: "${expression}" with data keys:`, Object.keys(data));
    
    try {
        // Handle simple property access first
        if (/^[a-zA-Z_][a-zA-Z0-9_.$\[\]]*$/.test(expression.trim())) {
            const result = this.getNestedValue(data, expression.trim());
            console.log(`Simple property access result:`, result);
            return result;
        }
        
        // Handle complex expressions
        return this.evaluateComplexExpression(expression, data);
    } catch (error) {
        console.error(`Expression evaluation error for "${expression}":`, error);
        return null;
    }
}

    evaluateCondition(condition, data) {
    console.log(`Evaluating condition: "${condition}"`);
    console.log('Available data:', data);
    
    try {
        // Normalize the condition
        const normalizedCondition = this.normalizeCondition(condition);
        console.log(`Normalized condition: "${normalizedCondition}"`);
        
        // Special handling for 'this' context in conditions
        let processedCondition = normalizedCondition;
        
        // Replace this.property with actual values
        processedCondition = processedCondition.replace(/\bthis\.([a-zA-Z_][a-zA-Z0-9_]*)\b/g, (match, property) => {
            if (data.this && data.this[property] !== undefined) {
                const value = data.this[property];
                const serialized = this.serializeValue(value);
                console.log(`Replacing ${match} with ${serialized}`);
                return serialized;
            }
            console.log(`Property ${property} not found in this context`);
            return 'null';
        });
        
        console.log(`Final condition for evaluation: "${processedCondition}"`);
        
        const result = this.evaluateComplexExpression(processedCondition, data);
        console.log(`Condition result: ${result}`);
        
        return result;
    } catch (error) {
        console.error(`Condition evaluation error for "${condition}":`, error);
        return false;
    }
}

    evaluateComplexExpression(expression, data) {
    console.log(`Evaluating complex expression: "${expression}"`);
    
    // Replace variable references with actual values
    let processedExpression = expression;
    
    // Find all variable references - enhanced pattern
    const variablePattern = /\b([a-zA-Z_][a-zA-Z0-9_.$\[\]]*)\b/g;
    const matches = [...expression.matchAll(variablePattern)];
    
    // Sort by length (longest first) to avoid partial replacements
    matches.sort((a, b) => b[0].length - a[0].length);
    
    for (const match of matches) {
        const variable = match[0];
        
        // Skip JavaScript keywords and operators
        if (this.isJavaScriptKeyword(variable)) continue;
        
        console.log(`Processing variable in expression: ${variable}`);
        
        const value = this.getNestedValue(data, variable);
        const serializedValue = this.serializeValue(value);
        
        console.log(`Replacing ${variable} with ${serializedValue}`);
        
        processedExpression = processedExpression.replace(
            new RegExp('\\b' + this.escapeRegExp(variable) + '\\b', 'g'),
            serializedValue
        );
    }
    
    console.log(`Processed expression: "${processedExpression}"`);
    
    // Safely evaluate the expression
    return this.safeEvaluate(processedExpression);
}

    normalizeCondition(condition) {
        // Handle common condition patterns
        return condition
            .replace(/\band\b/g, '&&')
            .replace(/\bor\b/g, '||')
            .replace(/\bnot\b/g, '!')
            .replace(/\bis\b/g, '==')
            .replace(/\bisn't\b/g, '!=')
            .replace(/\bequals\b/g, '==');
    }

    getNestedProperty(obj, path) {
    if (!obj || !path) return null;
    
    console.log(`Getting nested property "${path}" from:`, obj);
    
    // Handle array access like "deals[0].amount"
    const normalizedPath = path.replace(/\[(\d+)\]/g, '.$1');
    
    const keys = normalizedPath.split('.');
    let current = obj;
    
    for (const key of keys) {
        if (current === null || current === undefined) {
            console.log(`Null/undefined at key: ${key}`);
            return null;
        }
        
        // Handle array indices
        if (/^\d+$/.test(key)) {
            const index = parseInt(key);
            if (Array.isArray(current) && index >= 0 && index < current.length) {
                current = current[index];
                console.log(`Array access [${index}]:`, current);
            } else {
                console.log(`Array index ${index} out of bounds`);
                return null;
            }
        } else {
            if (current[key] !== undefined) {
                current = current[key];
                console.log(`Property .${key}:`, current);
            } else {
                console.log(`Property "${key}" not found`);
                return null;
            }
        }
    }
    
    console.log(`Final nested property value:`, current);
    return current;
}

    getNestedValue(obj, path) {
    if (!obj || !path) return null;
    
    console.log(`Getting nested value for path: "${path}" from object:`, obj);
    
    // Special handling for 'this.property' patterns
    if (path.startsWith('this.')) {
        const propertyPath = path.substring(5); // Remove 'this.'
        if (obj.this) {
            console.log(`Accessing this.${propertyPath} from:`, obj.this);
            return this.getNestedProperty(obj.this, propertyPath);
        } else {
            console.log('No "this" object found in context');
            return null;
        }
    }
    
    // Handle array access like users[0] or users[0].name
    const normalizedPath = path.replace(/\[(\d+)\]/g, '.$1');
    
    const keys = normalizedPath.split('.');
    let current = obj;
    
    for (let i = 0; i < keys.length; i++) {
        const key = keys[i];
        
        if (current === null || current === undefined) {
            console.log(`Null/undefined encountered at key: ${key}`);
            return null;
        }
        
        if (key === 'this') {
            // Access the 'this' property if it exists
            if (current.this !== undefined) {
                current = current.this;
                console.log(`Accessing 'this' context:`, current);
            } else {
                console.log('No "this" property found');
                return null;
            }
            continue;
        }
        
        // Handle array indices
        if (/^\d+$/.test(key)) {
            const index = parseInt(key);
            if (Array.isArray(current) && index >= 0 && index < current.length) {
                current = current[index];
                console.log(`Array access [${index}]:`, current);
            } else {
                console.log(`Array index ${index} out of bounds or not an array`);
                return null;
            }
        } else {
            if (current[key] !== undefined) {
                current = current[key];
                console.log(`Property access .${key}:`, current);
            } else {
                console.log(`Property "${key}" not found in object`);
                return null;
            }
        }
    }
    
    console.log(`Final value for "${path}":`, current);
    return current;
}


    serializeValue(value) {
        if (value === null || value === undefined) {
            return 'null';
        }
        
        if (typeof value === 'string') {
            return `"${value.replace(/"/g, '\\"')}"`;
        }
        
        if (typeof value === 'number' || typeof value === 'boolean') {
            return String(value);
        }
        
        if (Array.isArray(value)) {
            return JSON.stringify(value);
        }
        
        if (typeof value === 'object') {
            return JSON.stringify(value);
        }
        
        return `"${String(value)}"`;
    }

    safeEvaluate(expression) {
        // Create a safe evaluation context
        const allowedGlobals = {
            Math: Math,
            Date: Date,
            parseInt: parseInt,
            parseFloat: parseFloat,
            isNaN: isNaN,
            isFinite: isFinite
        };
        
        try {
            // Use Function constructor for safer evaluation
            const func = new Function(...Object.keys(allowedGlobals), `return (${expression});`);
            return func(...Object.values(allowedGlobals));
        } catch (error) {
            console.warn(`Safe evaluation failed for "${expression}":`, error.message);
            return null;
        }
    }

    isJavaScriptKeyword(word) {
        const keywords = [
            'true', 'false', 'null', 'undefined',
            'Math', 'Date', 'parseInt', 'parseFloat',
            'isNaN', 'isFinite', 'typeof', 'instanceof'
        ];
        return keywords.includes(word);
    }

    escapeRegExp(string) {
        return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
}

module.exports = ExpressionEvaluator;