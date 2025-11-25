# CLAUDE.md - AI Assistant Guide for 採用参謀AI (Recruitment Advisor AI)

**Project:** Google Apps Script MVP for Recruitment Intelligence System
**Version:** Phase 4-1 (Week 7)
**Last Updated:** 2025-11-25
**Technology Stack:** Google Apps Script (GAS), Google Sheets, Dify API Integration

---

## Table of Contents

1. [Project Overview](#project-overview)
2. [Codebase Structure](#codebase-structure)
3. [Key Architecture Concepts](#key-architecture-concepts)
4. [Development Workflows](#development-workflows)
5. [Coding Conventions](#coding-conventions)
6. [File Organization & Naming](#file-organization--naming)
7. [Core Systems & Components](#core-systems--components)
8. [Data Flow & Integration](#data-flow--integration)
9. [Testing & Validation](#testing--validation)
10. [Common Patterns & Best Practices](#common-patterns--best-practices)
11. [Troubleshooting Guide](#troubleshooting-guide)
12. [Deployment & Git Workflow](#deployment--git-workflow)

---

## Project Overview

### Purpose
採用参謀AI (Recruitment Advisor AI) is a Google Apps Script-based recruitment intelligence system that automates candidate evaluation, survey management, and acceptance probability calculation. The system helps recruitment teams:

- Track candidate progress through hiring stages
- Automatically calculate acceptance probability using AI-powered scoring
- Send and analyze survey responses at key interview phases
- Generate visual dashboards for data-driven decision making
- Integrate with Dify API for advanced AI capabilities

### System Architecture
```
┌─────────────────────────────────────────────────────────────┐
│                    Google Spreadsheet                        │
│  ┌────────────┬──────────────┬─────────────┬──────────────┐ │
│  │ Candidates │ Evaluation   │ Engagement  │ Dashboard    │ │
│  │ Master     │ Log          │ Log         │ (Charts)     │ │
│  └────────────┴──────────────┴─────────────┴──────────────┘ │
│                                                               │
│  ┌────────────────────────────────────────────────────────┐ │
│  │            Google Apps Script (Server-Side)            │ │
│  │  • Automated Survey Sending                            │ │
│  │  • Response Speed Calculation                          │ │
│  │  • Acceptance Probability Calculation                  │ │
│  │  • Dashboard Generation                                │ │
│  │  • Dify API Integration                                │ │
│  └────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────┘
         │                │               │
         ▼                ▼               ▼
   Gmail API      Google Forms    Dify API (Webhook)
   (Surveys)      (Responses)     (AI Analysis)
```

### Key Features
- **Automated Survey Management**: Sends timed surveys to candidates based on interview phase
- **Acceptance Probability Engine**: Multi-factor scoring algorithm with configurable weights
- **Real-time Dashboards**: Auto-updating charts and KPIs using QUERY functions
- **Time-based Triggers**: Hourly automated checks for new survey responses
- **Error Logging & Recovery**: Comprehensive error handling with retry mechanisms
- **Test Data Generation**: Automated test data creation for development

---

## Codebase Structure

### Root Directory Layout
```
gas-project_mvp_-/
├── Core Scripts (メインロジック)
│   ├── Code.gs (メインスクリプト).js          # Main entry point, custom menus
│   ├── Config.gs (設定値).js                  # Configuration constants
│   ├── AcceptanceCalculation.gs               # Acceptance probability calculations
│   ├── AutomatedCheck.gs                      # Time-based automated checks
│   └── ErrorHandler.gs                        # Error logging and recovery
│
├── Sheet Management (シート管理)
│   ├── SheetSetup.gs (シート作成・設定).js    # Sheet creation and setup
│   ├── DataValidation.gs (データ検証).js      # Data validation rules
│   ├── ConditionalFormatting.gs (条件付き書式).js  # Conditional formatting
│   └── InitialData.gs (初期データ投入).js     # Initial data population
│
├── Dashboard (Phase 4)
│   ├── DashboardSetup.gs                      # Dashboard sheet creation
│   ├── DashboardCharts.gs                     # Chart generation
│   ├── DASHBOARD_README.md                    # Dashboard documentation
│   └── DASHBOARD_IMPLEMENTATION_GUIDE.md      # Implementation guide
│
├── Data Operations (データ操作)
│   ├── DataGetters.gs (データ取得).js         # Data retrieval functions
│   ├── SpreadsheetExport.gs (スプレッドシート出力).js  # Export utilities
│   └── Utils.gs (ユーティリティ関数).js       # General utilities
│
├── Communication (通信)
│   ├── EmailSender.gs (アンケート送信).js     # Email/survey sending
│   └── DifyIntegration.gs (Dify連携).js       # Dify API integration
│
├── Testing (テスト)
│   ├── TestDataGenerator.gs                   # Test data generation
│   ├── TestHelpers.gs (テストヘルパー).js     # Test utilities
│   ├── TestPhase31.gs                         # Phase 3.1 tests
│   └── DiagnosticTest.gs                      # Diagnostic tests
│
├── Configuration
│   ├── appsscript.json                        # GAS project config
│   ├── .clasp.json                            # clasp deployment config
│   └── Triggers.gs (トリガー).js              # Trigger setup
│
└── Documentation
    ├── CLAUDE.md                              # This file
    ├── OPERATION_MANUAL.md                    # Operations manual
    ├── DASHBOARD_README.md                    # Dashboard guide
    └── DASHBOARD_IMPLEMENTATION_GUIDE.md      # Dashboard setup
```

### File Naming Conventions

**Pattern 1: Japanese with Function Description**
- Format: `{English}.gs ({Japanese}).js`
- Example: `Code.gs (メインスクリプト).js`
- Used for: Core user-facing components

**Pattern 2: English Descriptive Names**
- Format: `{PascalCase}.gs`
- Example: `AcceptanceCalculation.gs`
- Used for: Internal logic components

**Important**: When creating new files, follow the existing pattern in the directory. Core files use Pattern 1, logic files use Pattern 2.

---

## Key Architecture Concepts

### 1. Single Source of Truth (SSOT)
The system follows a strict SSOT principle:

- **Engagement_Log**: Master record of all candidate interactions and acceptance scores
- **Candidates_Master**: Current state aggregated via QUERY functions (read-only computed)
- **Dashboard_Data**: Intermediate aggregation layer for dashboard visualization

**Rule**: NEVER manually edit computed sheets. Always update source data (Engagement_Log).

### 2. Configuration-Driven Design
All constants are centralized in `Config.gs (設定値).js`:

```javascript
const CONFIG = {
  SHEET_NAMES: { ... },      // Sheet name constants
  COLORS: { ... },           // UI color schemes
  VALIDATION_OPTIONS: { ... }, // Dropdown options
  COLUMNS: { ... },          // Column index constants
  EMAIL: { ... },            // Email limits
  DIFY: { ... }              // API settings
};
```

**Rule**: ALWAYS use `CONFIG` constants instead of hardcoded strings or numbers.

### 3. Column Index Constants
Column positions are defined as zero-based constants in `CONFIG.COLUMNS`:

```javascript
CONFIG.COLUMNS.CANDIDATES_MASTER.CANDIDATE_ID = 0;  // A列
CONFIG.COLUMNS.CANDIDATES_MASTER.NAME = 1;          // B列
CONFIG.COLUMNS.CANDIDATES_MASTER.STATUS = 2;        // C列
```

**Rule**: Use column constants, never magic numbers: `row[CONFIG.COLUMNS.CANDIDATES_MASTER.NAME]`

### 4. Error Handling Strategy
Three-tier error handling:

1. **Logging**: All errors logged to Error_Log sheet via `logError()`
2. **Retry**: Network operations use `retryOnError()` with exponential backoff
3. **Recovery**: Processing_Log tracks state to prevent duplicate processing

**Pattern**:
```javascript
try {
  retryOnError(() => {
    // Risky operation
  }, 3, 2000);
} catch (error) {
  logError('functionName', error);
  throw error;
}
```

### 5. Time-Based Automation
The system uses time-driven triggers (hourly) via `AutomatedCheck.gs`:

```
┌─────────────────────────────────────────────┐
│  Time Trigger (1 hour)                      │
│  checkForNewSurveyResponses()               │
└──────────────┬──────────────────────────────┘
               │
               ▼
┌─────────────────────────────────────────────┐
│  Processing_Log Check                       │
│  - Get last processed row per phase         │
│  - Process only new rows                    │
│  - Update last processed row                │
└──────────────┬──────────────────────────────┘
               │
               ▼
┌─────────────────────────────────────────────┐
│  Engagement_Log Update                      │
│  - Calculate acceptance probability         │
│  - Log to Engagement_Log                    │
└─────────────────────────────────────────────┘
```

**Rule**: NEVER process all data. Always use Processing_Log to track incremental processing.

---

## Development Workflows

### Setting Up a New Development Session

1. **Check Current Branch**
   ```bash
   git status
   git log --oneline -5
   ```

2. **Review Recent Changes**
   - Read `OPERATION_MANUAL.md` for operations context
   - Check `Processing_Log` sheet for automation status
   - Review `Error_Log` sheet for active issues

3. **Understand Data State**
   - Check if test data exists: Menu > 🧪 テストデータ > データ状況を確認
   - Understand which phase you're working on (check recent commits)

### Making Code Changes

#### When Modifying Core Logic

1. **Read existing implementation first**
   ```javascript
   // Example: Before modifying acceptance calculation
   // Read: AcceptanceCalculation.gs to understand current logic
   ```

2. **Update Config.gs if adding new constants**
   ```javascript
   // Add to appropriate section
   CONFIG.COLUMNS.NEW_SHEET = {
     COLUMN_NAME: 0,
     // ...
   };
   ```

3. **Maintain column index mappings**
   - If adding columns to sheets, update `CONFIG.COLUMNS`
   - Update corresponding getter functions in `DataGetters.gs`

4. **Test with test data**
   ```javascript
   // Generate test data first
   generateAllTestData();

   // Then test your function
   yourNewFunction();
   ```

#### When Creating New Features

**Phase-based Development Pattern**:
The project follows a phase-based development model (currently Phase 4-1):

```
Phase 1: Basic sheet setup and manual data entry
Phase 2: Survey automation and response tracking
Phase 3: Acceptance calculation engine
Phase 4: Dashboard and visualization
  └─ Phase 4-1: In-spreadsheet dashboard (CURRENT)
```

When adding features:
1. Document which phase the feature belongs to
2. Update version comments in affected files
3. Add to appropriate documentation (OPERATION_MANUAL.md or specific README)

### Testing Workflow

1. **Unit Testing**: Test individual functions
   ```javascript
   function testYourFunction() {
     const result = yourFunction(testInput);
     Logger.log(`Expected: X, Got: ${result}`);
   }
   ```

2. **Integration Testing**: Use `TestPhase31.gs` pattern
   ```javascript
   function testYourIntegration() {
     // Setup test data
     const candidateId = 'TEST001';

     // Execute workflow
     sendSurvey(candidateId);
     processResponse(candidateId);

     // Verify results
     const score = getAcceptanceScore(candidateId);
     Logger.log(`Score: ${score}`);
   }
   ```

3. **Data Validation**: Use automated validators
   ```javascript
   validateAllTestData(); // Checks data structure
   checkTestDataStatus(); // Checks data existence
   ```

### Debugging Workflow

1. **Use Logger.log() extensively**
   ```javascript
   Logger.log(`Processing candidate: ${candidateId}`);
   Logger.log(`Score calculation: base=${base}, bonus=${bonus}, total=${total}`);
   ```

2. **Check Error_Log sheet**
   - All errors automatically logged with stack traces
   - Filter by function name or date

3. **Use Diagnostic Functions**
   ```javascript
   checkSurveyData();      // Verify survey data structure
   debugContactHistory();  // Debug contact history issues
   ```

4. **Execution Logs**
   - Apps Script Editor > View > Executions
   - Shows all trigger executions and errors

---

## Coding Conventions

### JavaScript Style

```javascript
/**
 * Function documentation with JSDoc
 *
 * @param {string} candidateId - Candidate unique ID
 * @param {string} phase - Interview phase
 * @return {number} Calculated acceptance score
 */
function calculateAcceptanceScore(candidateId, phase) {
  // 1. Validate inputs
  if (!candidateId || !phase) {
    throw new Error('Missing required parameters');
  }

  // 2. Get data using constants
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAMES.ENGAGEMENT_LOG);

  // 3. Process with error handling
  try {
    const score = performCalculation(candidateId, phase);
    return score;
  } catch (error) {
    logError('calculateAcceptanceScore', error);
    throw error;
  }
}
```

### Naming Conventions

- **Functions**: `camelCase`, descriptive verbs
  - Good: `calculateAcceptanceScore()`, `sendSurveyEmail()`
  - Bad: `calc()`, `send()`

- **Constants**: `UPPER_SNAKE_CASE`
  - Good: `DEFAULT_SCORES.MOTIVATION`
  - Bad: `defaultMotivation`

- **Variables**: `camelCase`, descriptive nouns
  - Good: `candidateData`, `surveyResponse`
  - Bad: `data`, `resp`

- **Sheet Names**: Use `CONFIG.SHEET_NAMES` constants
  - Good: `CONFIG.SHEET_NAMES.CANDIDATES_MASTER`
  - Bad: `'Candidates_Master'` (hardcoded string)

### Comment Style

**Japanese + English for clarity**:
```javascript
// ========== Phase 3-1: 基礎要素スコア計算 ==========
// Foundation Score Calculation

/**
 * 志望度スコアを計算
 * Calculate motivation score based on survey response
 */
function calculateMotivationScore(candidateId, phase) {
  // ...
}
```

**Section Headers**:
```javascript
// ========================================
// Phase 3-1: 基礎要素スコア計算
// ========================================
```

### Data Access Patterns

**Always use helper functions from DataGetters.gs**:
```javascript
// Good: Use centralized getter
const email = getCandidateEmail(candidateId);

// Bad: Direct sheet access
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Candidates_Master');
const data = sheet.getDataRange().getValues();
// ... manual search ...
```

**Query Pattern**:
```javascript
// Use QUERY functions for aggregation (in sheets, not in scripts)
=QUERY(Engagement_Log!A:U,
  "SELECT B, MAX(F)
   WHERE B IS NOT NULL
   GROUP BY B
   LABEL MAX(F) '最新承諾可能性'")
```

---

## File Organization & Naming

### Core Script Files

| File | Purpose | Key Functions |
|------|---------|---------------|
| `Code.gs (メインスクリプト).js` | Entry point, custom menus | `onOpen()`, `setupAllSheets()` |
| `Config.gs (設定値).js` | All configuration constants | `CONFIG` object |
| `AcceptanceCalculation.gs` | Acceptance probability engine | `calculateFoundationScore()` |
| `AutomatedCheck.gs` | Time-based automation | `checkForNewSurveyResponses()` |
| `ErrorHandler.gs` | Error management | `logError()`, `retryOnError()` |

### Sheet Management Files

| File | Purpose | Key Functions |
|------|---------|---------------|
| `SheetSetup.gs (シート作成・設定).js` | Create and configure sheets | `createAllSheets()`, `setupCandidatesMaster()` |
| `DataValidation.gs (データ検証).js` | Data validation rules | `setupAllDataValidation()` |
| `ConditionalFormatting.gs (条件付き書式).js` | Visual formatting rules | `setupAllConditionalFormatting()` |
| `InitialData.gs (初期データ投入).js` | Sample data population | `insertAllInitialData()` |

### Dashboard Files (Phase 4-1)

| File | Purpose | Key Functions |
|------|---------|---------------|
| `DashboardSetup.gs` | Dashboard sheet creation | `setupDashboardPhase4()` |
| `DashboardCharts.gs` | Chart generation | `createAcceptanceChart()` |

### Helper Files

| File | Purpose | Key Functions |
|------|---------|---------------|
| `DataGetters.gs (データ取得).js` | Data retrieval utilities | `getCandidateEmail()`, `getSurveyResponse()` |
| `Utils.gs (ユーティリティ関数).js` | General utilities | `formatDate()`, `sanitizeString()` |
| `EmailSender.gs (アンケート送信).js` | Email automation | `sendSurveyEmail()` |
| `DifyIntegration.gs (Dify連携).js` | Dify API integration | `callDifyAPI()`, `doPost()` webhook |

---

## Core Systems & Components

### 1. Acceptance Probability Calculation Engine

**Location**: `AcceptanceCalculation.gs`

**Formula Structure**:
```
【初回面談・社員面談】
承諾可能性 = 基礎要素(40%) + 関係性要素(30%) + 行動シグナル要素(30%)

【2次面接・内定後】
承諾可能性 = 基礎要素(40%) + 関係性要素(30%) + 行動シグナル要素(20%) + 自己申告要素(10%)
```

**Key Functions**:
- `calculateFoundationScore()`: Base score from survey responses
  - Motivation score (志望度)
  - Competitive advantage (競合優位性)
  - Concern resolution (懸念解消度)
- `calculateRelationshipScore()`: Relationship indicators (future implementation)
- `calculateBehavioralScore()`: Action signals (future implementation)
- `calculateSelfReportedScore()`: Self-reported confidence (2次面接+ only)

**Default Values** (from `DEFAULT_SCORES`):
```javascript
const DEFAULT_SCORES = {
  MOTIVATION: 50,
  COMPETITIVE: 50,
  CONCERN_RESOLUTION: 70,
  FOUNDATION: 50
};
```

**When to modify**:
- Adjusting weight percentages: Update calculation functions
- Adding new factors: Add to appropriate score category
- Changing defaults: Update `DEFAULT_SCORES` object

### 2. Automated Survey System

**Location**: `EmailSender.gs (アンケート送信).js` + `AutomatedCheck.gs`

**Survey Phases**:
1. 初回面談 (First Interview)
2. 社員面談 (Employee Interview)
3. 2次面接 (Second Interview)
4. 内定後 (Post-Offer)

**Survey URLs** (configured in `Config.gs`):
```javascript
CONFIG.SURVEY_URLS = {
  '初回面談': 'https://docs.google.com/forms/d/...',
  '社員面談': 'https://docs.google.com/forms/d/...',
  '2次面接': 'https://docs.google.com/forms/d/...',
  '内定後': 'https://docs.google.com/forms/d/...'
};
```

**Workflow**:
```
1. Manual send via menu → sendSurveyEmail()
2. Email sent via Gmail API → Logged to Survey_Send_Log
3. Candidate responds → Google Forms submission
4. Hourly trigger → checkForNewSurveyResponses()
5. New responses processed → calculateAndLogAcceptance()
6. Results logged → Engagement_Log
```

**Important Constraints**:
- Gmail limit: 90 emails/day (configured in `CONFIG.EMAIL.DAILY_LIMIT`)
- Retry: 3 attempts with 2-second delay
- Deduplication: Processing_Log tracks last processed row

### 3. Dashboard System (Phase 4-1)

**Location**: `DashboardSetup.gs`, `DashboardCharts.gs`

**Architecture**:
```
Engagement_Log (Source)
    │
    ├─ QUERY functions ─→ Dashboard_Data (Intermediate)
    │                           │
    │                           ├─ Latest Acceptance (A1:E100)
    │                           ├─ Phase Score Trend (G1:K100)
    │                           ├─ AI vs Human (M1:P100)
    │                           ├─ Concern Analysis (R1:S20)
    │                           └─ Status Distribution (U1:V5)
    │
    └─ Referenced by ──────→ Dashboard (Main Sheet)
                                 │
                                 ├─ KPI Summary
                                 ├─ Candidate Ranking
                                 ├─ AI vs Human Comparison
                                 └─ 4 Charts (bar, line, scatter, pie)
```

**Setup Function**:
```javascript
setupDashboardPhase4(); // One-time setup
```

**Key Features**:
- Auto-updating via QUERY functions (no manual refresh)
- Conditional formatting (red/yellow/green for scores)
- 4 chart types for different analyses
- KPIs: Total candidates, average score, high-probability count

**Customization Points**:
- Add KPI: Edit `setupDashboardKPIs()`
- Change colors: Edit `DashboardCharts.gs` color options
- Adjust thresholds: Edit `setupDashboardConditionalFormats()`

### 4. Processing State Management

**Location**: `AutomatedCheck.gs`

**Processing_Log Structure**:
```
| Phase      | Last Processed Row | Last Updated      |
|------------|-------------------|-------------------|
| 初回面談    | 5                 | 2025-11-25 10:00  |
| 社員面談    | 3                 | 2025-11-25 10:00  |
| 2次面接     | 7                 | 2025-11-25 10:00  |
| 内定後      | 2                 | 2025-11-25 10:00  |
```

**Functions**:
- `getLastProcessedRow(phase)`: Get checkpoint
- `updateLastProcessedRow(phase, rowIndex)`: Update checkpoint
- `resetProcessingLog()`: Clear all checkpoints (re-process all data)

**Important**: Processing_Log prevents duplicate entries in Engagement_Log.

### 5. Error Management System

**Location**: `ErrorHandler.gs`

**Error_Log Structure**:
```
| Timestamp | Function Name | Error Message | Error Stack | Status |
```

**Functions**:
- `logError(functionName, error)`: Log to Error_Log sheet
- `retryOnError(func, maxRetries, delay)`: Retry with exponential backoff
- `validateData(candidateId, phase)`: Pre-execution validation

**Pattern**:
```javascript
try {
  retryOnError(() => {
    // Network call or risky operation
    sendSurveyEmail(candidateId, phase);
  }, CONFIG.EMAIL.RETRY_COUNT, CONFIG.EMAIL.RETRY_DELAY);
} catch (error) {
  logError('sendSurveyEmail', error);
  // Decide: throw or continue with defaults
}
```

---

## Data Flow & Integration

### Primary Data Flow

```
┌─────────────────────────────────────────────────────────────┐
│ 1. Manual Data Entry (Candidates_Master, Contact_History)   │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ 2. Survey Sent (Menu → sendSurveyEmail)                     │
│    Logged to: Survey_Send_Log                               │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ 3. Candidate Responds (Google Forms)                        │
│    Data appears in: Survey sheets (IMPORTRANGE)             │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ 4. Hourly Trigger (checkForNewSurveyResponses)              │
│    - Check Processing_Log                                   │
│    - Get new rows since last checkpoint                     │
│    - Process each new response                              │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ 5. Calculate Acceptance (AcceptanceCalculation.gs)          │
│    - Get survey response data                               │
│    - Calculate foundation score                             │
│    - Calculate final acceptance probability                 │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ 6. Log to Engagement_Log (Single Source of Truth)           │
│    - candidate_id, acceptance_rate_final, phase, etc.       │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ 7. Auto-Update Dashboard (QUERY functions)                  │
│    - Dashboard_Data refreshes                               │
│    - Dashboard shows updated KPIs and charts                │
└─────────────────────────────────────────────────────────────┘
```

### Sheet Dependencies

**Read-Only Computed Sheets** (never manually edit):
- `Candidates_Master`: Aggregates from Engagement_Log via QUERY
- `Dashboard_Data`: Aggregates from Engagement_Log via QUERY
- `Dashboard`: References Dashboard_Data

**Source Sheets** (manually edited or script-updated):
- `Engagement_Log`: **Primary source** for all acceptance data
- `Evaluation_Log`: Interview evaluation scores
- `Contact_History`: Communication log
- `Survey_Send_Log`: Survey sending log
- Survey response sheets (via IMPORTRANGE)

### External Integrations

#### Dify API Integration
**File**: `DifyIntegration.gs (Dify連携).js`

**Configuration**:
```javascript
// Stored in Script Properties (not in code)
PropertiesService.getScriptProperties().setProperty('DIFY_API_URL', 'https://...');
PropertiesService.getScriptProperties().setProperty('DIFY_API_KEY', 'app-...');
```

**Setup**:
```javascript
setupDifyApiSettings(); // Via menu
```

**Webhook Endpoint**:
```javascript
function doPost(e) {
  // Receives webhook from Dify
  // Processes payload
  // Updates appropriate sheets
}
```

**Get Webhook URL**:
```javascript
showWebhookUrl(); // Via menu
```

#### Google Forms Integration
**Pattern**: IMPORTRANGE to pull form responses into the spreadsheet

**Survey Sheets** (one per phase):
- 初回面談_Survey
- 社員面談_Survey
- 2次面接_Survey
- 内定後_Survey

**Important**: Email addresses in forms MUST match `Candidates_Master.EMAIL` for proper candidate linking.

---

## Testing & Validation

### Test Data Generation

**Main Function**: `generateAllTestData()` in `TestDataGenerator.gs`

**Generates**:
- 15 test candidates in Candidates_Master
- Evaluation_Log entries
- Engagement_Log entries
- Survey_Send_Log entries
- Survey response data
- Contact_History entries

**Usage**:
```javascript
// Via menu
Menu > 🧪 テストデータ > すべてのテストデータを生成

// Or via script
generateAllTestData();
```

**Clear Test Data**:
```javascript
clearAllTestData(); // Removes all test data
```

**Check Data Status**:
```javascript
checkTestDataStatus(); // Shows row counts for all sheets
```

### Validation Functions

**Structure Validation**:
```javascript
validateAllTestData(); // Comprehensive structure check
validateSheetColumnCount(sheetName); // Check column count
validateTestDataStructure(); // Check data integrity
```

**Data Validators**:
```javascript
validateData(candidateId, phase); // Validate before processing
```

### Test Patterns

**Unit Test Pattern**:
```javascript
function testMotivationScore() {
  const candidateId = 'TEST001';
  const phase = '初回面談';

  const score = calculateMotivationScore(candidateId, phase);

  Logger.log(`Motivation Score: ${score}`);
  Logger.log(`Expected: 50-100, Got: ${score}`);

  if (score < 0 || score > 100) {
    throw new Error(`Invalid score: ${score}`);
  }
}
```

**Integration Test Pattern**:
```javascript
function testSurveyWorkflow() {
  // 1. Setup
  generateAllTestData();

  // 2. Send survey
  const candidateId = 'TEST001';
  sendSurveyEmail(candidateId, '初回面談');

  // 3. Simulate response processing
  checkForNewSurveyResponses();

  // 4. Verify results
  const score = getLatestAcceptanceScore(candidateId);
  Logger.log(`Final acceptance score: ${score}`);

  // 5. Cleanup
  clearTestData();
}
```

### Diagnostic Tools

```javascript
// Check survey data structure
checkSurveyData();

// Debug contact history
debugContactHistory();

// Test automated check system
testAutomatedCheckSystem();
```

---

## Common Patterns & Best Practices

### Pattern 1: Safe Sheet Access

```javascript
// Bad: No error handling
const sheet = SpreadsheetApp.getActiveSpreadsheet()
  .getSheetByName('SomeName');
const data = sheet.getDataRange().getValues();

// Good: With error handling
function getSheetData(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }

    const data = sheet.getDataRange().getValues();
    if (data.length === 0) {
      Logger.log(`Warning: Sheet ${sheetName} is empty`);
      return [];
    }

    return data;
  } catch (error) {
    logError('getSheetData', error);
    throw error;
  }
}
```

### Pattern 2: Column Access with Constants

```javascript
// Bad: Magic numbers
const candidateId = row[0];
const name = row[1];
const status = row[2];

// Good: Using constants
const COL = CONFIG.COLUMNS.CANDIDATES_MASTER;
const candidateId = row[COL.CANDIDATE_ID];
const name = row[COL.NAME];
const status = row[COL.STATUS];
```

### Pattern 3: Date Handling

```javascript
// Good: Consistent date formatting
function formatDate(date) {
  if (!(date instanceof Date)) {
    date = new Date(date);
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

// Usage
const now = formatDate(new Date());
```

### Pattern 4: Email Sending with Limits

```javascript
function sendEmailWithLimit(to, subject, body) {
  // Check daily limit
  const todayCount = getTodaySendCount();
  if (todayCount >= CONFIG.EMAIL.DAILY_LIMIT) {
    throw new Error('Daily email limit reached');
  }

  // Send with retry
  retryOnError(() => {
    GmailApp.sendEmail(to, subject, body);
  }, CONFIG.EMAIL.RETRY_COUNT, CONFIG.EMAIL.RETRY_DELAY);

  // Log send
  logSurveyEmail(to, subject, 'success');
}
```

### Pattern 5: Processing with Checkpoints

```javascript
function processNewData(phase) {
  const sheet = getSheetByPhase(phase);
  const data = sheet.getDataRange().getValues();

  // Get checkpoint
  const lastRow = getLastProcessedRow(phase);
  const newRows = data.slice(lastRow + 1);

  Logger.log(`Processing ${newRows.length} new rows for ${phase}`);

  newRows.forEach((row, index) => {
    const currentRow = lastRow + index + 1;

    try {
      processRow(row);

      // Update checkpoint after each row
      updateLastProcessedRow(phase, currentRow);
    } catch (error) {
      logError('processNewData', error);
      // Continue with next row or stop depending on error severity
    }
  });
}
```

### Pattern 6: QUERY Function Usage (in Sheets)

```javascript
// In sheet setup scripts, use QUERY for aggregation
function setupLatestScores() {
  const formula = `=QUERY(Engagement_Log!A:U,
    "SELECT B, MAX(F)
     WHERE B IS NOT NULL
     GROUP BY B
     ORDER BY MAX(F) DESC
     LABEL B '候補者ID', MAX(F) '最新承諾可能性'")`;

  sheet.getRange('A1').setFormula(formula);
}
```

**QUERY Best Practices**:
- Always filter NULL: `WHERE B IS NOT NULL`
- Use LABEL for Japanese headers
- Limit results: `LIMIT 100`
- Use appropriate aggregate: MAX, MIN, AVG, COUNT

### Pattern 7: Default Value Handling

```javascript
function getScoreWithDefault(candidateId, scoreType) {
  try {
    const score = getScore(candidateId, scoreType);

    // Check if valid score
    if (score === null || score === undefined || isNaN(score)) {
      Logger.log(`Using default for ${scoreType}: ${DEFAULT_SCORES[scoreType]}`);
      return DEFAULT_SCORES[scoreType];
    }

    return score;
  } catch (error) {
    Logger.log(`Error getting ${scoreType}, using default: ${error.message}`);
    return DEFAULT_SCORES[scoreType];
  }
}
```

---

## Troubleshooting Guide

### Common Issues and Solutions

#### Issue 1: New Survey Responses Not Appearing in Engagement_Log

**Symptoms**:
- Survey responses visible in survey sheets
- No new entries in Engagement_Log
- Processing_Log not updating

**Diagnosis**:
1. Check Processing_Log "Last Updated" timestamp
2. Check Apps Script Triggers (menu: Apps Script > Triggers)
3. Check Error_Log for recent errors

**Solutions**:
```javascript
// Solution 1: Check trigger exists
// Go to: Apps Script > Triggers
// Verify: checkForNewSurveyResponses is set to run hourly

// Solution 2: Manual trigger
checkForNewSurveyResponses();

// Solution 3: Reset and reprocess
resetProcessingLog();
checkForNewSurveyResponses();
```

#### Issue 2: Duplicate Entries in Engagement_Log

**Symptoms**:
- Same candidate/phase appears multiple times
- Processing_Log not incrementing

**Diagnosis**:
1. Check for multiple triggers (Apps Script > Triggers)
2. Check Processing_Log update logic

**Solutions**:
```javascript
// Solution 1: Remove duplicate triggers
// Go to: Apps Script > Triggers
// Delete all but one checkForNewSurveyResponses trigger

// Solution 2: Clear duplicates and reset
// Manually delete duplicate rows in Engagement_Log
resetProcessingLog();
```

#### Issue 3: Acceptance Scores Always Show Default Values

**Symptoms**:
- All scores are 50 (default)
- Survey data exists but not used

**Diagnosis**:
1. Check email matching: Survey email must match Candidates_Master
2. Check survey sheet column structure
3. Check getters in DataGetters.gs

**Solutions**:
```javascript
// Solution 1: Verify email match
// Compare: Survey sheet email column vs Candidates_Master EMAIL column

// Solution 2: Debug specific candidate
const email = getCandidateEmail('TEST001');
Logger.log(`Candidate email: ${email}`);

const surveyData = getSurveyResponse(email, '初回面談');
Logger.log(`Survey data: ${JSON.stringify(surveyData)}`);

// Solution 3: Check column indices in Config.gs
```

#### Issue 4: Dashboard Not Updating

**Symptoms**:
- Dashboard shows old data or errors
- QUERY functions show #REF! or #ERROR!

**Diagnosis**:
1. Check Engagement_Log has data
2. Check QUERY formula syntax
3. Check Dashboard_Data sheet exists

**Solutions**:
```javascript
// Solution 1: Regenerate dashboard
setupDashboardPhase4();

// Solution 2: Check QUERY formulas
// Open Dashboard_Data sheet
// Check each QUERY formula for errors

// Solution 3: Force recalculation
// Edit > Find and replace
// Find: = Replace: = (triggers recalc)
```

#### Issue 5: Script Execution Timeout

**Symptoms**:
- "Exceeded maximum execution time" error
- Long-running functions fail

**Diagnosis**:
1. Check how many rows being processed
2. Check for inefficient loops

**Solutions**:
```javascript
// Solution 1: Process in batches
function processBatch(startRow, batchSize) {
  const endRow = Math.min(startRow + batchSize, maxRow);
  // Process rows startRow to endRow
}

// Solution 2: Use time-based chunking
function processWithTimeLimit(maxSeconds = 300) {
  const startTime = new Date();

  while (hasMoreWork()) {
    const elapsed = (new Date() - startTime) / 1000;
    if (elapsed > maxSeconds) {
      Logger.log('Time limit reached, stopping');
      break;
    }

    processNextItem();
  }
}
```

### Error Log Analysis

**Common Error Messages**:

| Error | Cause | Solution |
|-------|-------|----------|
| `Sheet not found: X` | Sheet doesn't exist or wrong name | Verify sheet exists, check CONFIG.SHEET_NAMES |
| `Cannot read property 'getRange' of null` | Sheet is null | Add null check before accessing sheet |
| `Service invoked too many times` | API quota exceeded | Add delays between API calls |
| `Exception: Invalid email` | Email format wrong | Validate email before sending |
| `TypeError: Cannot read property '0' of undefined` | Empty array access | Check array length before accessing |

**Error Log Query**:
```javascript
// Get all unresolved errors
function getUnresolvedErrors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Error_Log');
  const data = sheet.getDataRange().getValues();

  const unresolved = data.filter(row => row[4] === 'Unresolved');
  Logger.log(`Unresolved errors: ${unresolved.length}`);
  unresolved.forEach(row => {
    Logger.log(`${row[0]} | ${row[1]} | ${row[2]}`);
  });
}
```

---

## Deployment & Git Workflow

### Branch Naming Convention

**Pattern**: `claude/{description}-{session-id}`

**Example**: `claude/claude-md-mie2867ygr49na9w-01NPuYAjScGo23RG9ZgbTs2S`

**Important**:
- Branch MUST start with `claude/`
- Branch MUST end with matching session ID
- Otherwise push will fail with 403 HTTP error

### Git Commit Guidelines

**Commit Message Format**:
```
<type>(<scope>): <description>

[optional body]
```

**Types**:
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation changes
- `refactor`: Code refactoring
- `test`: Test additions/changes
- `chore`: Maintenance tasks

**Examples**:
```
feat(dashboard): Add Phase 4-1 in-spreadsheet dashboard

- Created DashboardSetup.gs and DashboardCharts.gs
- Implemented auto-updating QUERY functions
- Added 4 chart types (bar, line, scatter, pie)

fix(automation): Handle getUi() error in script editor context

- Wrapped getUi() calls in try-catch
- Script now works in both UI and editor contexts
```

### Git Push with Retry

**Always use the designated branch**:
```bash
# Check current branch
git status

# Push to the correct branch
git push -u origin claude/claude-md-mie2867ygr49na9w-01NPuYAjScGo23RG9ZgbTs2S
```

**If network errors occur**, retry up to 4 times with exponential backoff:
```bash
# Retry sequence (automated in CI/CD)
# Attempt 1: immediate
# Attempt 2: wait 2s
# Attempt 3: wait 4s
# Attempt 4: wait 8s
# Attempt 5: wait 16s
```

### Pre-commit Checklist

Before committing:

- [ ] Code tested with test data (`generateAllTestData()`)
- [ ] No hardcoded values (use CONFIG constants)
- [ ] Error handling added for risky operations
- [ ] Logger.log statements added for debugging
- [ ] Column constants updated if new columns added
- [ ] Documentation updated (if public-facing feature)
- [ ] No sensitive data (API keys, emails) in code

### Deployment to Apps Script

**Using clasp** (recommended):
```bash
# Login to Google
clasp login

# Push to Apps Script
clasp push

# View logs
clasp logs
```

**Manual Deployment**:
1. Open Apps Script editor
2. Copy-paste changed files
3. Save (Ctrl+S)
4. Test in spreadsheet

**Important**: `.clasp.json` contains script ID - do not modify unless creating new project.

### Creating Pull Requests

**When ready to merge**:

1. **Verify all changes work**:
   ```javascript
   // Run comprehensive test
   validateAllTestData();
   setupDashboardPhase4(); // If dashboard changes
   ```

2. **Create PR with detailed description**:
   ```markdown
   ## Summary
   - Implemented Phase 4-1 dashboard
   - Added auto-updating QUERY functions
   - Created 4 visualization charts

   ## Test Plan
   - [x] Generated test data
   - [x] Verified dashboard updates automatically
   - [x] Checked all QUERY formulas work
   - [x] Tested conditional formatting

   ## Screenshots
   [Include dashboard screenshot]
   ```

3. **Use gh CLI** (if available):
   ```bash
   gh pr create --title "feat(dashboard): Phase 4-1 implementation" --body "$(cat <<'EOF'
   ## Summary
   ...
   EOF
   )"
   ```

---

## Key Principles for AI Assistants

### When Working on This Codebase

1. **Always Read Before Modifying**
   - Read existing implementation before suggesting changes
   - Understand the phase-based development model
   - Check recent commits to understand context

2. **Maintain Single Source of Truth**
   - Engagement_Log is the master record
   - Never manually edit computed sheets (Candidates_Master, Dashboard)
   - Always log new data to source sheets

3. **Use Configuration Constants**
   - Never hardcode sheet names, use `CONFIG.SHEET_NAMES`
   - Never hardcode column indices, use `CONFIG.COLUMNS`
   - Never hardcode limits, use `CONFIG.EMAIL`, `CONFIG.DIFY`

4. **Preserve Error Handling**
   - All risky operations wrapped in try-catch
   - Use `retryOnError()` for network calls
   - Always call `logError()` for debugging

5. **Follow Japanese Naming Conventions**
   - Core files use Japanese descriptions
   - Comments use Japanese + English
   - Respect existing bilingual patterns

6. **Test with Test Data**
   - Always generate test data before testing
   - Use diagnostic functions to verify
   - Clean up test data after testing

7. **Document Phase Changes**
   - Mark new features with phase number
   - Update version comments
   - Add to appropriate documentation

8. **Respect Time Limits**
   - Apps Script has 6-minute execution limit
   - Use batching for long operations
   - Update Processing_Log frequently

9. **Maintain Backwards Compatibility**
   - Don't break existing QUERY formulas
   - Don't rename sheets without updating all references
   - Don't change column order without updating CONFIG.COLUMNS

10. **Follow Git Workflow**
    - Use correct branch naming (`claude/*-{session-id}`)
    - Write descriptive commit messages
    - Test before committing

---

## Quick Reference

### Essential Functions

```javascript
// Setup & Initialization
setupAllSheets()                    // Complete system setup
setupDashboardPhase4()              // Dashboard setup (Phase 4-1)

// Data Generation
generateAllTestData()               // Generate test data
clearAllTestData()                  // Clear test data
checkTestDataStatus()               // Check data status

// Automation
checkForNewSurveyResponses()        // Process new survey responses
resetProcessingLog()                // Reset processing checkpoints

// Survey Management
sendSurveyEmail(id, phase)          // Send survey to candidate

// Validation
validateAllTestData()               // Validate data structure
validateData(candidateId, phase)    // Validate specific data

// Debugging
checkSurveyData()                   // Debug survey data
debugContactHistory()               // Debug contact history
getUnresolvedErrors()               // Get error log entries

// Error Handling
logError(functionName, error)       // Log error to Error_Log
retryOnError(func, retries, delay)  // Retry with backoff
```

### Key Configuration Paths

```javascript
// Sheet names
CONFIG.SHEET_NAMES.ENGAGEMENT_LOG
CONFIG.SHEET_NAMES.CANDIDATES_MASTER
CONFIG.SHEET_NAMES.DASHBOARD

// Column indices
CONFIG.COLUMNS.CANDIDATES_MASTER.CANDIDATE_ID
CONFIG.COLUMNS.ENGAGEMENT_LOG.LOG_ID

// Colors
CONFIG.COLORS.HEADER_BG
CONFIG.COLORS.HIGH_SCORE

// Limits
CONFIG.EMAIL.DAILY_LIMIT
CONFIG.DIFY.TIMEOUT
```

### Important Sheets

| Sheet | Type | Purpose |
|-------|------|---------|
| Engagement_Log | Source | Master acceptance data (SSOT) |
| Candidates_Master | Computed | Aggregated candidate state |
| Dashboard_Data | Computed | Dashboard intermediate data |
| Dashboard | Computed | Visual dashboard |
| Processing_Log | State | Processing checkpoints |
| Error_Log | Log | Error tracking |
| Survey_Send_Log | Log | Email sending log |

---

## Version History

| Date | Version | Changes |
|------|---------|---------|
| 2025-11-25 | Phase 4-1 | Dashboard implementation complete |
| 2025-11-19 | Phase 3.5 | Automated survey processing |
| 2025-11-12 | Phase 3.1 | Acceptance calculation engine |

---

## Additional Resources

- **Operations Manual**: See `OPERATION_MANUAL.md` for daily/weekly operations
- **Dashboard Guide**: See `DASHBOARD_README.md` for dashboard usage
- **Implementation Guide**: See `DASHBOARD_IMPLEMENTATION_GUIDE.md` for setup
- **Apps Script Docs**: https://developers.google.com/apps-script
- **Clasp Documentation**: https://github.com/google/clasp

---

**End of CLAUDE.md** - This document should be updated whenever significant architectural changes are made to the codebase.
