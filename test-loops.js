// Updated test-loops.js with correct processing order
const TemplateEngine = require('./src/core/TemplateEngine');

async function testLoopsCorrectOrder() {
    console.log('Testing Loop Processing with Correct Order...\n');
    
    const engine = new TemplateEngine();
    
    const testData = {
        salesTeam: [
            {
                name: "Sarah Mitchell",
                region: "North America", 
                target: 300000,
                actual: 345000,
                achievement: 1.15,
                status: "exceeding"
            },
            {
                name: "Michael Chen",
                region: "Asia Pacific",
                target: 250000, 
                actual: 215000,
                achievement: 0.86,
                status: "behind target"
            }
        ]
    };
    
    // Simple test XML with loop
    const testXml = `
        <w:document>
            <w:body>
                <w:p><w:t>Sales Team Performance</w:t></w:p>
                \${#each salesTeam}
                <w:p><w:t>\${this.name} - \${this.region} Region</w:t></w:p>
                <w:p><w:t>Target: \${this.target|currency}</w:t></w:p>
                <w:p><w:t>Achievement: \${this.achievement|percent}</w:t></w:p>
                <w:p><w:t>Status: \${#if this.achievement >= 1.0}\${this.status|upper}\${#else}\${this.status}\${/if}</w:t></w:p>
                \${/each}
            </w:body>
        </w:document>
    `;
    
    console.log('Input XML:');
    console.log(testXml);
    console.log('\nTest Data:');
    console.log(JSON.stringify(testData, null, 2));
    
    try {
        // Test the NEW processing pipeline (loops first!)
        console.log('\n--- Using CORRECT processing order ---');
        const result = await engine.processAdvancedTemplate(Buffer.from('dummy'), testData);
        
        console.log('\nResult should contain proper values now...');
        
        // For testing, let's call the individual methods in correct order
        console.log('\n--- Manual step-by-step processing ---');
        
        let step1 = engine.processLoops(testXml, testData);
        console.log('\n1. After processLoops:');
        console.log(step1.substring(0, 500) + '...');
        
        let step2 = engine.processConditions(step1, testData);  
        console.log('\n2. After processConditions:');
        console.log(step2.substring(0, 500) + '...');
        
        let step3 = engine.processRemainingVariables(step2, testData);
        console.log('\n3. After processRemainingVariables:');
        console.log(step3.substring(0, 500) + '...');
        
        // Validation
        const containsNames = step3.includes('Sarah Mitchell') && step3.includes('Michael Chen');
        const containsRegions = step3.includes('North America') && step3.includes('Asia Pacific');
        const containsNumbers = step3.includes('300000') && step3.includes('215000');
        const containsFormatting = step3.includes('$') || step3.includes('%');
        
        console.log('\nValidation:');
        console.log(`Contains names: ${containsNames}`);
        console.log(`Contains regions: ${containsRegions}`);
        console.log(`Contains numbers: ${containsNumbers}`); 
        console.log(`Contains formatting: ${containsFormatting}`);
        
        const success = containsNames && containsRegions && containsNumbers;
        console.log(`\n${success ? 'PASSED' : 'FAILED'}: Loop processing test`);
        
        return success;
        
    } catch (error) {
        console.error('Test failed with error:', error);
        return false;
    }
}

// Simple direct test of the processLoops method
async function testLoopsDirectly() {
    console.log('\n--- Direct Loop Processing Test ---');
    
    const engine = new TemplateEngine();
    
    const simpleData = {
        users: [
            { name: "Alice", score: 95 },
            { name: "Bob", score: 87 }
        ]
    };
    
    const simpleXml = `
        Users:
        \${#each users}
        - \${this.name}: \${this.score} points
        \${/each}
    `;
    
    console.log('Simple test data:', JSON.stringify(simpleData, null, 2));
    console.log('Simple XML:', simpleXml);
    
    try {
        const result = engine.processLoops(simpleXml, simpleData);
        console.log('\nDirect loop result:');
        console.log(result);
        
        const hasAlice = result.includes('Alice');
        const hasBob = result.includes('Bob'); 
        const hasScores = result.includes('95') && result.includes('87');
        
        console.log(`\nDirect test: ${hasAlice && hasBob && hasScores ? 'PASSED' : 'FAILED'}`);
        
        return hasAlice && hasBob && hasScores;
        
    } catch (error) {
        console.error('Direct loop test failed:', error);
        return false;
    }
}

// Run both tests
if (require.main === module) {
    testLoopsDirectly()
        .then((directResult) => {
            console.log(`\nDirect loop test: ${directResult ? 'PASSED' : 'FAILED'}`);
            return testLoopsCorrectOrder();
        })
        .then((fullResult) => {
            console.log(`\nFull processing test: ${fullResult ? 'PASSED' : 'FAILED'}`);
            console.log('\n' + '='.repeat(50));
            console.log('Overall result: Both tests should PASS after the fix!');
        })
        .catch(console.error);
}

module.exports = { testLoopsCorrectOrder, testLoopsDirectly };