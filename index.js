// Load all dependencies
require("dotenv").config();
const express = require("express");
const multer = require("multer");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const pdfParse = require("pdf-parse");
const mammoth = require("mammoth");
const {
  GoogleGenerativeAI,
  HarmBlockThreshold,
  HarmCategory,
} = require("@google/generative-ai");
const xlsx = require("xlsx");

const app = express();
app.use(cors());
app.use(express.json());

// Configure multer for file uploads
const upload = multer({
  dest: "uploads/",
  limits: { fileSize: 25 * 1024 * 1024 },
});

// Initialize Gemini AI client
const genAI = new GoogleGenerativeAI(process.env.API_KEY);

// ========== PRECISION-FOCUSED CONFIGURATIONS ==========

// Enhanced JSON schema for maximum precision in extraction
const PRECISION_EXTRACTION_SCHEMA = {
  type: "object",
  properties: {
    metadata: {
      type: "object",
      properties: {
        patientName: { type: "string", nullable: true },
        age: { type: "string", nullable: true },
        gender: { 
          type: "string", 
          nullable: true,
          enum: [null, "male", "female", "both", "unknown"]
        },
        dateOfBirth: { type: "string", nullable: true },
        reportGeneratedDate: { type: "string", nullable: true },
        labName: { type: "string", nullable: true },
        doctorName: { type: "string", nullable: true },
        reportId: { type: "string", nullable: true }
      },
      required: []
    },
    biomarkers: {
      type: "array",
      items: {
        type: "object",
        properties: {
          name: { 
            type: "string",
            description: "Exact name of the biomarker as it appears in the report"
          },
          alternativeNames: {
            type: "array",
            items: { type: "string" },
            description: "Any alternative names or abbreviations found for this biomarker"
          },
          group: { 
            type: "string", 
            nullable: true,
            description: "Test panel or group this biomarker belongs to (e.g., CBC, CMP, Lipid Panel)"
          },
          unit: { 
            type: "string", 
            nullable: true,
            description: "Exact unit of measurement as written in the report"
          },
          value: { 
            type: "string", 
            nullable: true,
            description: "Exact value as it appears, including any symbols like <, >, or ranges"
          },
          numericValue: {
            type: "number",
            nullable: true,
            description: "Parsed numeric value if the value can be converted to a number"
          },
          referenceRange: {
            type: "string",
            nullable: true,
            description: "Reference range as stated in the report, if available"
          },
          flag: {
            type: "string",
            nullable: true,
            enum: [null, "HIGH", "LOW", "NORMAL", "ABNORMAL", "CRITICAL", "H", "L", "N", "*"],
            description: "Any flags or indicators next to the result"
          },
          method: {
            type: "string",
            nullable: true,
            description: "Testing method if mentioned"
          },
          specimen: {
            type: "string",
            nullable: true,
            description: "Specimen type (serum, plasma, whole blood, urine, etc.)"
          },
          confidence: {
            type: "number",
            minimum: 0.0,
            maximum: 1.0,
            description: "Confidence score for the accuracy of this extraction (0.0 to 1.0)"
          }
        },
        required: ["name", "confidence"]
      }
    },
    extraction_metadata: {
      type: "object",
      properties: {
        total_biomarkers_found: { type: "number" },
        text_quality_score: { 
          type: "number",
          minimum: 0.0,
          maximum: 1.0,
          description: "Assessment of how clear and readable the input text was"
        },
        extraction_challenges: {
          type: "array",
          items: { type: "string" },
          description: "List any challenges encountered during extraction"
        }
      },
      required: ["total_biomarkers_found", "text_quality_score"]
    }
  },
  required: ["metadata", "biomarkers", "extraction_metadata"]
};

// Ultra-precise extraction prompt with detailed instructions
const ULTRA_PRECISE_EXTRACTION_PROMPT = `
You are an expert medical laboratory technologist with 20+ years of experience in interpreting lab reports. Your task is to extract biomarker data with maximum precision and accuracy.

CRITICAL PRECISION REQUIREMENTS:

1. **EXACT VALUE EXTRACTION**:
   - Copy values EXACTLY as they appear: "14.5", "<0.5", ">100", "1.2-3.4"
   - Never round, estimate, or modify values
   - Include all symbols: <, >, ±, ~, etc.
   - For ranges like "1.2-3.4", extract as single value string

2. **UNIT PRECISION**:
   - Copy units EXACTLY: "mg/dL", "μg/L", "ng/mL", "IU/L", "mmol/L"
   - Don't standardize or convert units
   - Include all special characters: μ, ², ³, /, etc.
   - Watch for temperature units: °C, °F

3. **NAME STANDARDIZATION RULES**:
   - Use the most complete name available in the report
   - If multiple names exist, use the primary one for 'name' field
   - List all variations in 'alternativeNames' array
   - Common standardizations:
     * "Hemoglobin" (not "Hgb", "Hb", "HEMOGLOBIN")
     * "White Blood Cell Count" (not "WBC", "Leukocytes")
     * "Total Cholesterol" (not "CHOL", "Cholesterol")
     * "Glucose" (not "GLU", "Blood Sugar")

4. **FLAG INTERPRETATION**:
   - Extract flags exactly: "H", "L", "HIGH", "LOW", "*", "CRITICAL"
   - Don't interpret or translate flags
   - Include any custom lab flags

5. **CONFIDENCE SCORING**:
   - 1.0: Perfect extraction, clear text, unambiguous values
   - 0.9: Very good extraction, minor formatting issues
   - 0.8: Good extraction, some text quality issues
   - 0.7: Acceptable extraction, moderate challenges
   - 0.6: Poor extraction, significant text issues
   - <0.6: Very poor extraction, major uncertainties

6. **QUALITY ASSESSMENT**:
   - Assess overall text readability (OCR quality, formatting)
   - Note any extraction challenges encountered
   - Count total biomarkers found

7. **COMPREHENSIVE EXTRACTION**:
   Extract ALL test results including:
   - Complete Blood Count (CBC) with differential
   - Basic/Comprehensive Metabolic Panel (BMP/CMP)
   - Lipid profiles
   - Liver function tests (LFTs)
   - Kidney function tests
   - Cardiac markers (Troponins, CK-MB, BNP)
   - Inflammatory markers (CRP, ESR)
   - Thyroid function (TSH, T3, T4)
   - Diabetes markers (HbA1c, Glucose)
   - Tumor markers
   - Coagulation studies (PT, PTT, INR)
   - Urinalysis components
   - Immunology/Serology results
   - Microbiology results
   - Any other laboratory values

EXAMPLES OF PRECISE EXTRACTION:

Input: "Glucose: 126 mg/dL (H) [Normal: 70-99 mg/dL]"
Output: {
  "name": "Glucose",
  "value": "126",
  "numericValue": 126,
  "unit": "mg/dL",
  "flag": "H",
  "referenceRange": "70-99 mg/dL",
  "confidence": 1.0
}

Input: "WBC (White Blood Cells): <0.5 x10³/μL"
Output: {
  "name": "White Blood Cell Count",
  "alternativeNames": ["WBC", "White Blood Cells"],
  "value": "<0.5",
  "numericValue": null,
  "unit": "x10³/μL",
  "confidence": 1.0
}

SOURCE REPORT:
<<<
{{RAW_TEXT}}
>>>
`.trim();

// Enhanced analysis schema with precision metrics
const PRECISION_ANALYSIS_SCHEMA = {
  type: "object",
  properties: {
    evaluated: {
      type: "array",
      items: {
        type: "object",
        properties: {
          name: { type: "string" },
          originalValue: { type: "string", nullable: true },
          numericValue: { type: "number", nullable: true },
          unit: { type: "string", nullable: true },
          matched_marker: { type: "string", nullable: true },
          matching_confidence: {
            type: "number",
            minimum: 0.0,
            maximum: 1.0,
            description: "Confidence in the biomarker matching (0.0 to 1.0)"
          },
          unit_match_exact: {
            type: "boolean",
            description: "Whether the units matched exactly with dataset"
          },
          gender_used: { type: "string", nullable: true },
          min_reference: { type: "string", nullable: true },
          max_reference: { type: "string", nullable: true },
          min_numeric: { type: "number", nullable: true },
          max_numeric: { type: "number", nullable: true },
          severity_score: { type: "string", nullable: true },
          isWithinRange: { type: "boolean", nullable: true },
          deviation_percentage: {
            type: "number",
            nullable: true,
            description: "Percentage deviation from normal range midpoint"
          },
          clinical_significance: {
            type: "string",
            nullable: true,
            enum: [null, "normal", "borderline", "abnormal", "critical"],
            description: "Clinical significance assessment"
          },
          interpretation_notes: {
            type: "string",
            nullable: true,
            description: "Additional clinical interpretation notes"
          }
        },
        required: ["name", "matching_confidence", "unit_match_exact"]
      }
    },
    unmatched: {
      type: "array",
      items: {
        type: "object",
        properties: {
          name: { type: "string" },
          originalValue: { type: "string", nullable: true },
          unit: { type: "string", nullable: true },
          unmatch_reason: {
            type: "string",
            enum: [
              "no_marker_match",
              "unit_mismatch", 
              "gender_mismatch",
              "insufficient_reference_data",
              "qualitative_result",
              "rare_biomarker"
            ],
            description: "Specific reason why this biomarker couldn't be matched"
          },
          similar_markers: {
            type: "array",
            items: { type: "string" },
            description: "List of similar markers found in dataset"
          },
          confidence: {
            type: "number",
            minimum: 0.0,
            maximum: 1.0,
            description: "Confidence that this biomarker extraction is correct"
          }
        },
        required: ["name", "unmatch_reason", "confidence"]
      }
    },
    precision_metrics: {
      type: "object",
      properties: {
        total_biomarkers: { type: "number" },
        successfully_matched: { type: "number" },
        exact_unit_matches: { type: "number" },
        approximate_unit_matches: { type: "number" },
        gender_specific_matches: { type: "number" },
        both_gender_matches: { type: "number" },
        average_matching_confidence: {
          type: "number",
          minimum: 0.0,
          maximum: 1.0
        },
        critical_values_count: { type: "number" },
        abnormal_values_count: { type: "number" },
        normal_values_count: { type: "number" },
        matching_precision_score: {
          type: "number",
          minimum: 0.0,
          maximum: 1.0,
          description: "Overall precision score for the matching process"
        }
      },
      required: [
        "total_biomarkers", "successfully_matched", "exact_unit_matches",
        "average_matching_confidence", "matching_precision_score"
      ]
    }
  },
  required: ["evaluated", "unmatched", "precision_metrics"]
};

// Ultra-precise analysis prompt
const ULTRA_PRECISE_ANALYSIS_PROMPT = `
You are a precision-focused clinical laboratory data analyst. Your task is to match biomarkers with maximum accuracy and provide detailed precision metrics.

PRECISION MATCHING PROTOCOL:

1. **BIOMARKER NAME MATCHING**:
   - Use medical knowledge for synonym matching
   - Score matching confidence based on name similarity
   - Common matches to recognize:
     * Hemoglobin = Hgb = Hb = HGB
     * White Blood Cell Count = WBC = Leukocytes = White Blood Cells
     * Total Cholesterol = Cholesterol = CHOL = TC
     * Alanine Aminotransferase = ALT = SGPT
     * Aspartate Aminotransferase = AST = SGOT
     * Blood Urea Nitrogen = BUN = Urea Nitrogen
     * Creatinine = Cr = CREAT
     * Glucose = GLU = Blood Sugar = BG

2. **UNIT MATCHING PRECISION**:
   - EXACT match required for high confidence
   - Flag unit mismatches clearly
   - Common equivalent units to recognize:
     * mg/dL = mg/dl = mg/100mL
     * μg/L = mcg/L = ug/L
     * mIU/L = mU/L (for some hormones)
     * x10³/μL = K/μL = 10³/μL
   - NEVER match incompatible units (e.g., mg/dL vs mmol/L without conversion)

3. **GENDER MATCHING HIERARCHY**:
   - Priority 1: Exact gender match ("male" for male patient)
   - Priority 2: "both" gender entry
   - Priority 3: No match (mark as unmatched)
   - NEVER use wrong gender reference ranges

4. **NUMERICAL ANALYSIS PRECISION**:
   - Parse numeric values carefully: "14.5", "<10", ">100"
   - Handle comparison operators correctly
   - Calculate precise deviation percentages
   - Identify critical vs abnormal vs borderline values

5. **CONFIDENCE SCORING SYSTEM**:
   - 1.0: Perfect name match + exact unit match + appropriate gender
   - 0.9: Perfect name match + exact unit match + "both" gender
   - 0.8: Strong synonym match + exact unit match
   - 0.7: Good name match + compatible unit match
   - 0.6: Moderate name match + unit issues
   - <0.6: Poor matching confidence

6. **UNMATCH CATEGORIZATION**:
   - "no_marker_match": No similar biomarker found in dataset
   - "unit_mismatch": Biomarker found but units incompatible
   - "gender_mismatch": Found marker but wrong gender reference
   - "insufficient_reference_data": Marker found but missing min/max
   - "qualitative_result": Result is text-based (Positive/Negative)
   - "rare_biomarker": Uncommon test not in reference dataset

7. **CLINICAL SIGNIFICANCE ASSESSMENT**:
   - "normal": Within reference range
   - "borderline": Just outside range (<20% deviation)
   - "abnormal": Significantly outside range (20-100% deviation)
   - "critical": Dangerously outside range (>100% deviation)

=== PATIENT DATA ===
Gender: {{GENDER}}

=== EXTRACTED BIOMARKERS ===
{{BIOMARKERS}}

=== REFERENCE DATASET ===
{{DATASET_CSV}}

IMPORTANT: Provide detailed precision metrics and explain your matching decisions with confidence scores.
`.trim();

// ========== PRECISION HELPER FUNCTIONS ==========

/**
 * Enhanced text preprocessing for better OCR/extraction accuracy
 */
function preprocessTextForPrecision(rawText) {
  // Remove multiple spaces but preserve structure
  let processed = rawText.replace(/[ \t]{2,}/g, ' ');
  
  // Fix common OCR issues
  processed = processed
    // Fix common character substitutions
    .replace(/[Oo](?=\d)/g, '0')  // O -> 0 before digits
    .replace(/[Il](?=\d)/g, '1')  // I,l -> 1 before digits
    .replace(/(\d)[Oo](\d)/g, '$10$2')  // 0 between digits
    .replace(/(\d)[Il](\d)/g, '$11$2')  // 1 between digits
    
    // Fix unit issues
    .replace(/mg\/d[Il]/g, 'mg/dL')  // Common mg/dL OCR error
    .replace(/μg\/[Il]/g, 'μg/L')    // Common μg/L OCR error
    .replace(/mcg\/[Il]/g, 'mcg/L')  // Common mcg/L OCR error
    
    // Fix common lab value patterns
    .replace(/([A-Za-z]+)\s*[:]\s*([<>]?\s*\d+\.?\d*)/g, '$1: $2')
    
    // Normalize whitespace around colons and values
    .replace(/\s*:\s*/g, ': ')
    .replace(/([<>])\s+(\d)/g, '$1$2')  // Remove space after < or >
    
    // Clean up extra whitespace
    .replace(/\n{3,}/g, '\n\n')
    .trim();
    
  return processed;
}

/**
 * Advanced biomarker name standardization
 */
function standardizeBiomarkerName(name) {
  const standardizations = {
    // Hematology
    'wbc': 'White Blood Cell Count',
    'white blood cells': 'White Blood Cell Count',
    'leukocytes': 'White Blood Cell Count',
    'rbc': 'Red Blood Cell Count',
    'red blood cells': 'Red Blood Cell Count',
    'hgb': 'Hemoglobin',
    'hb': 'Hemoglobin',
    'hct': 'Hematocrit',
    'plt': 'Platelet Count',
    'platelets': 'Platelet Count',
    'mcv': 'Mean Corpuscular Volume',
    'mch': 'Mean Corpuscular Hemoglobin',
    'mchc': 'Mean Corpuscular Hemoglobin Concentration',
    'rdw': 'Red Cell Distribution Width',
    
    // Chemistry
    'glu': 'Glucose',
    'glucose': 'Glucose',
    'blood sugar': 'Glucose',
    'bun': 'Blood Urea Nitrogen',
    'cr': 'Creatinine',
    'creat': 'Creatinine',
    'na': 'Sodium',
    'k': 'Potassium',
    'cl': 'Chloride',
    'co2': 'Carbon Dioxide',
    'chol': 'Total Cholesterol',
    'cholesterol': 'Total Cholesterol',
    'hdl': 'HDL Cholesterol',
    'ldl': 'LDL Cholesterol',
    'trig': 'Triglycerides',
    'triglycerides': 'Triglycerides',
    
    // Liver function
    'alt': 'Alanine Aminotransferase',
    'sgpt': 'Alanine Aminotransferase',
    'ast': 'Aspartate Aminotransferase',
    'sgot': 'Aspartate Aminotransferase',
    'alkp': 'Alkaline Phosphatase',
    'alp': 'Alkaline Phosphatase',
    'tbili': 'Total Bilirubin',
    'bilirubin': 'Total Bilirubin',
    
    // Thyroid
    'tsh': 'Thyroid Stimulating Hormone',
    't4': 'Thyroxine',
    't3': 'Triiodothyronine',
    
    // Cardiac
    'ck': 'Creatine Kinase',
    'cpk': 'Creatine Kinase',
    'ck-mb': 'Creatine Kinase-MB',
    'troponin i': 'Troponin I',
    'troponin t': 'Troponin T',
    
    // Inflammatory
    'crp': 'C-Reactive Protein',
    'esr': 'Erythrocyte Sedimentation Rate',
    
    // Diabetes
    'hba1c': 'Hemoglobin A1c',
    'a1c': 'Hemoglobin A1c'
  };
  
  const lowerName = name.toLowerCase().trim();
  return standardizations[lowerName] || name;
}

/**
 * Load dataset with enhanced error handling and validation
 */
function loadDatasetWithPrecision(xlsxPath) {
  if (!fs.existsSync(xlsxPath)) {
    throw new Error(`Dataset file not found: ${xlsxPath}`);
  }
  
  const wb = xlsx.readFile(xlsxPath);
  const sheetName = wb.SheetNames[0];
  const worksheet = wb.Sheets[sheetName];
  
  // Convert to JSON with careful handling
  const rawData = xlsx.utils.sheet_to_json(worksheet, { 
    defval: null,
    raw: false, // Keep as strings to preserve precision
    dateNF: 'yyyy-mm-dd'
  });
  
  // Validate and clean dataset entries
  const cleanedData = rawData.filter(row => {
    const marker = row["Blood Test Marker"] || row["blood test marker"];
    return marker && marker.trim().length > 0;
  }).map(row => ({
    marker: (row["Blood Test Marker"] || row["blood test marker"] || "").trim(),
    gender: (row["Gender"] || row["gender"] || "").toLowerCase().trim(),
    min: row["Minimum"] || row["minimum"] || "",
    max: row["Maximum"] || row["maximum"] || "",
    severity: row["Severity Score (1 = mild, 5 = highly significant)"] || 
              row["severity score (1 = mild, 5 = highly significant)"] || ""
  }));
  
  console.log(`Loaded ${cleanedData.length} validated dataset entries`);
  return cleanedData;
}

/**
 * Generate high-precision CSV for analysis
 */
function generatePrecisionCSV(datasetRows) {
  const headers = ["Marker", "Gender", "Min", "Max", "Severity"];
  const csvLines = [headers.join(",")];
  
  datasetRows.forEach(row => {
    const csvRow = [
      `"${row.marker.replace(/"/g, '""')}"`,  // Escape quotes
      `"${row.gender}"`,
      `"${row.min}"`,
      `"${row.max}"`,
      `"${row.severity}"`
    ];
    csvLines.push(csvRow.join(","));
  });
  
  return csvLines.join("\n");
}

// ========== API ROUTES WITH PRECISION FOCUS ==========

app.post("/extractReportGemini", upload.single("document"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ success: false, error: "No file uploaded." });
  }
  
  const filePath = path.resolve(req.file.path);
  
  try {
    const buffer = fs.readFileSync(filePath);
    let rawText = "";

    // Extract text with precision handling
    if (req.file.mimetype === "application/pdf") {
      const parsed = await pdfParse(buffer);
      rawText = parsed.text;
    } else if (req.file.mimetype.includes("wordprocessingml")) {
      const result = await mammoth.extractRawText({ path: filePath });
      rawText = result.value;
    } else {
      throw new Error("Unsupported file type. Only PDF and DOCX supported.");
    }

    // Apply precision preprocessing
    rawText = preprocessTextForPrecision(rawText);

    // Check text quality
    if (rawText.length < 100) {
      return res.status(400).json({
        success: false,
        error: "Insufficient readable text extracted from document."
      });
    }

    // Truncate if too long (with precision preservation)
    const MAX_CHARS = 3_500_000;
    if (rawText.length > MAX_CHARS) {
      console.warn(`Text truncated from ${rawText.length} to ${MAX_CHARS} characters`);
      rawText = rawText.substring(0, MAX_CHARS);
    }

    // Get model and generate response
    const model = genAI.getGenerativeModel({
      model: process.env.AI_MODEL || "gemini-1.5-pro",
    });

    const prompt = ULTRA_PRECISE_EXTRACTION_PROMPT.replace("{{RAW_TEXT}}", rawText);

    const response = await model.generateContent({
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: {
        temperature: 0.1, // Lower temperature for higher precision
        maxOutputTokens: 65536,
        responseMimeType: "application/json",
        responseSchema: PRECISION_EXTRACTION_SCHEMA,
      },
      safetySettings: [
        {
          category: HarmCategory.HARM_CATEGORY_HARASSMENT,
          threshold: HarmBlockThreshold.BLOCK_NONE,
        },
        {
          category: HarmCategory.HARM_CATEGORY_HATE_SPEECH,
          threshold: HarmBlockThreshold.BLOCK_NONE,
        },
        {
          category: HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT,
          threshold: HarmBlockThreshold.BLOCK_NONE,
        },
        {
          category: HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
          threshold: HarmBlockThreshold.BLOCK_NONE,
        },
      ],
    });

    const rawResponse = response.response.text();
    const resultData = JSON.parse(rawResponse);

    // Post-process for additional precision improvements (minimal structure)
    if (resultData.biomarkers) {
      resultData.biomarkers = resultData.biomarkers.map(biomarker => {
        // Clean and standardize the biomarker data
        const cleanedBiomarker = {
          name: standardizeBiomarkerName(biomarker.name),
          group: biomarker.group || null,
          unit: biomarker.unit || null,
          value: biomarker.value || null,
          numericValue: null
        };

        // Parse numeric values with precision
        if (biomarker.value && !isNaN(parseFloat(biomarker.value.replace(/[<>≤≥]/g, '')))) {
          cleanedBiomarker.numericValue = parseFloat(biomarker.value.replace(/[<>≤≥]/g, ''));
        }

        return cleanedBiomarker;
      });
    }

    res.json({ success: true, data: resultData });

  } catch (err) {
    console.error("Precision extraction failed:", err.message);
    res.status(500).json({
      success: false,
      message: `Extraction failed: ${err.message}`,
      debug: process.env.NODE_ENV === 'development' ? err.stack : undefined
    });
  } finally {
    // Clean up file
    fs.unlink(filePath, (unlinkErr) => {
      if (unlinkErr) console.error("Error deleting file:", unlinkErr);
    });
  }
});

app.post("/analyzeWithDataset", async (req, res) => {
  try {
    const { biomarkers, metadata } = req.body.data;

    if (!Array.isArray(biomarkers)) {
      return res.status(400).json({
        success: false,
        message: "Request body must contain a 'data.biomarkers' array."
      });
    }

    const gender = (metadata?.gender || "").toLowerCase();
    
    // Load dataset with precision handling
    const datasetPath = path.join(__dirname, "All_Descriptions_Completed.xlsx");
    const dataset = loadDatasetWithPrecision(datasetPath);
    const datasetCSV = generatePrecisionCSV(dataset);

    // Enhanced biomarkers preprocessing
    const processedBiomarkers = biomarkers.map(biomarker => ({
      ...biomarker,
      name: standardizeBiomarkerName(biomarker.name || ''),
      // Ensure numeric parsing
      numericValue: biomarker.value && !isNaN(parseFloat(biomarker.value.replace(/[<>]/g, ''))) 
        ? parseFloat(biomarker.value.replace(/[<>]/g, '')) 
        : null
    }));

    const model = genAI.getGenerativeModel({
      model: process.env.AI_MODEL || "gemini-1.5-pro",
    });

    const prompt = ULTRA_PRECISE_ANALYSIS_PROMPT
      .replace("{{GENDER}}", gender)
      .replace("{{BIOMARKERS}}", JSON.stringify(processedBiomarkers, null, 2))
      .replace("{{DATASET_CSV}}", datasetCSV);

    const response = await model.generateContent({
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: {
        temperature: 0.1, // Ultra-low temperature for precision
        maxOutputTokens: 65536,
        responseMimeType: "application/json",
        responseSchema: PRECISION_ANALYSIS_SCHEMA,
      },
      safetySettings: [
        {
          category: HarmCategory.HARM_CATEGORY_HARASSMENT,
          threshold: HarmBlockThreshold.BLOCK_NONE,
        },
        {
          category: HarmCategory.HARM_CATEGORY_HATE_SPEECH,
          threshold: HarmBlockThreshold.BLOCK_NONE,
        },
        {
          category: HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT,
          threshold: HarmBlockThreshold.BLOCK_NONE,
        },
        {
          category: HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
          threshold: HarmBlockThreshold.BLOCK_NONE,
        },
      ],
    });

    const rawResponse = response.response.text();
    const analysisResult = JSON.parse(rawResponse);

    // Add precision validation metrics
    const validationMetrics = {
      extraction_precision: {
        input_biomarkers: biomarkers.length,
        processed_biomarkers: processedBiomarkers.length,
        standardized_names: processedBiomarkers.filter(b => 
          b.name !== biomarkers[processedBiomarkers.indexOf(b)]?.name
        ).length,
        numeric_values_parsed: processedBiomarkers.filter(b => b.numericValue !== null).length,
        units_present: processedBiomarkers.filter(b => b.unit && b.unit.trim()).length
      },
      matching_precision: {
        total_attempted: analysisResult.evaluated.length + analysisResult.unmatched.length,
        high_confidence_matches: analysisResult.evaluated.filter(e => e.matching_confidence >= 0.8).length,
        exact_unit_matches: analysisResult.evaluated.filter(e => e.unit_match_exact === true).length,
        gender_appropriate_matches: analysisResult.evaluated.filter(e => 
          e.gender_used === gender || e.gender_used === 'both'
        ).length
      },
      clinical_precision: {
        within_range_count: analysisResult.evaluated.filter(e => e.isWithinRange === true).length,
        outside_range_count: analysisResult.evaluated.filter(e => e.isWithinRange === false).length,
        critical_findings: analysisResult.evaluated.filter(e => 
          e.clinical_significance === 'critical'
        ).length,
        abnormal_findings: analysisResult.evaluated.filter(e => 
          e.clinical_significance === 'abnormal' || e.clinical_significance === 'critical'
        ).length
      }
    };

    // Enhanced result with precision metrics
    const finalResult = {
      ...analysisResult,
      validation_metrics: validationMetrics,
      processing_info: {
        dataset_entries: dataset.length,
        analysis_timestamp: new Date().toISOString(),
        model_temperature: 0.1,
        precision_mode: 'ultra-high'
      }
    };

    res.json({ success: true, ...finalResult });

  } catch (err) {
    console.error("Precision analysis failed:", err.message);
    
    if (err.message.includes("Unexpected end of JSON")) {
      res.status(500).json({
        success: false,
        message: "AI analysis was incomplete. This may indicate the response was too complex. Please try again.",
        error_type: "incomplete_json_response"
      });
    } else if (err.code === 'ENOENT') {
      res.status(500).json({
        success: false,
        message: "Dataset file not found. Please ensure 'All_Descriptions_Completed.xlsx' is available.",
        error_type: "missing_dataset"
      });
    } else {
      res.status(500).json({
        success: false,
        message: `Analysis failed: ${err.message}`,
        error_type: "analysis_error",
        debug: process.env.NODE_ENV === 'development' ? err.stack : undefined
      });
    }
  }
});

// ========== PRECISION VALIDATION ENDPOINTS ==========

// Endpoint to validate extraction precision
app.post("/validateExtraction", async (req, res) => {
  try {
    const { extractedData, originalText } = req.body;
    
    if (!extractedData || !originalText) {
      return res.status(400).json({
        success: false,
        message: "Both extractedData and originalText are required."
      });
    }

    // Perform precision validation checks
    const validationResults = {
      text_quality_checks: {
        original_length: originalText.length,
        readable_ratio: calculateReadabilityRatio(originalText),
        ocr_quality_score: assessOCRQuality(originalText),
        structure_preservation: assessStructurePreservation(originalText)
      },
      extraction_quality_checks: {
        biomarker_count: extractedData.biomarkers?.length || 0,
        values_with_units: extractedData.biomarkers?.filter(b => b.unit).length || 0,
        numeric_values: extractedData.biomarkers?.filter(b => b.numericValue !== null).length || 0,
        high_confidence_extractions: extractedData.biomarkers?.filter(b => b.confidence >= 0.8).length || 0,
        flagged_results: extractedData.biomarkers?.filter(b => b.flag).length || 0
      },
      precision_scores: {
        overall_extraction_score: calculateOverallExtractionScore(extractedData),
        value_precision_score: calculateValuePrecisionScore(extractedData.biomarkers || []),
        unit_precision_score: calculateUnitPrecisionScore(extractedData.biomarkers || []),
        name_standardization_score: calculateNameStandardizationScore(extractedData.biomarkers || [])
      },
      recommendations: generatePrecisionRecommendations(extractedData, originalText)
    };

    res.json({ success: true, validation: validationResults });

  } catch (err) {
    console.error("Validation failed:", err.message);
    res.status(500).json({
      success: false,
      message: `Validation failed: ${err.message}`
    });
  }
});

// Endpoint to get precision analytics
app.get("/precisionAnalytics", async (req, res) => {
  try {
    const datasetPath = path.join(__dirname, "All_Descriptions_Completed.xlsx");
    const dataset = loadDatasetWithPrecision(datasetPath);
    
    const analytics = {
      dataset_analytics: {
        total_markers: dataset.length,
        gender_distribution: {
          male: dataset.filter(d => d.gender === 'male').length,
          female: dataset.filter(d => d.gender === 'female').length,
          both: dataset.filter(d => d.gender === 'both').length,
          unknown: dataset.filter(d => !d.gender || d.gender === 'unknown').length
        },
        severity_distribution: {
          mild: dataset.filter(d => d.severity === '1').length,
          moderate: dataset.filter(d => d.severity === '2' || d.severity === '3').length,
          significant: dataset.filter(d => d.severity === '4' || d.severity === '5').length,
          unspecified: dataset.filter(d => !d.severity).length
        },
        reference_range_completeness: {
          both_min_max: dataset.filter(d => d.min && d.max).length,
          only_min: dataset.filter(d => d.min && !d.max).length,
          only_max: dataset.filter(d => !d.min && d.max).length,
          neither: dataset.filter(d => !d.min && !d.max).length
        }
      },
      common_biomarkers: getCommonBiomarkerStats(dataset),
      precision_capabilities: {
        supported_units: extractUniqueUnits(dataset),
        gender_specific_markers: dataset.filter(d => d.gender !== 'both').length,
        critical_markers: dataset.filter(d => d.severity === '5').length
      }
    };

    res.json({ success: true, analytics });

  } catch (err) {
    console.error("Analytics failed:", err.message);
    res.status(500).json({
      success: false,
      message: `Analytics failed: ${err.message}`
    });
  }
});

// ========== PRECISION HELPER FUNCTIONS ==========

function calculateReadabilityRatio(text) {
  const totalChars = text.length;
  const readableChars = text.replace(/[^\w\s\.\,\:\;\(\)\[\]\-\+\<\>\=\/]/g, '').length;
  return totalChars > 0 ? readableChars / totalChars : 0;
}

function assessOCRQuality(text) {
  let score = 1.0;
  
  // Check for common OCR errors
  const ocrErrorPatterns = [
    /[Il](?=\d)/g,  // I or l before digits
    /[Oo](?=\d)/g,  // O before digits
    /(\d)[Il]/g,    // I or l after digits
    /(\d)[Oo]/g,    // O after digits
    /[^\w\s\.\,\:\;\(\)\[\]\-\+\<\>\=\/\%\°\μ]/g  // Unusual characters
  ];
  
  ocrErrorPatterns.forEach(pattern => {
    const matches = (text.match(pattern) || []).length;
    score -= matches * 0.001; // Small penalty per error
  });
  
  return Math.max(0, Math.min(1, score));
}

function assessStructurePreservation(text) {
  // Look for typical lab report structure elements
  const structureElements = [
    /test\s*name|marker|parameter/i,
    /result|value/i,
    /reference|normal|range/i,
    /unit|measure/i,
    /flag|status/i
  ];
  
  const foundElements = structureElements.filter(pattern => pattern.test(text)).length;
  return foundElements / structureElements.length;
}

function calculateOverallExtractionScore(extractedData) {
  const biomarkers = extractedData.biomarkers || [];
  if (biomarkers.length === 0) return 0;
  
  const avgConfidence = biomarkers.reduce((sum, b) => sum + (b.confidence || 0), 0) / biomarkers.length;
  const completenessScore = (biomarkers.filter(b => b.value && b.unit).length / biomarkers.length);
  const textQualityScore = extractedData.extraction_metadata?.text_quality_score || 0.5;
  
  return (avgConfidence * 0.4 + completenessScore * 0.4 + textQualityScore * 0.2);
}

function calculateValuePrecisionScore(biomarkers) {
  if (biomarkers.length === 0) return 0;
  
  const valuesWithNumbers = biomarkers.filter(b => b.numericValue !== null).length;
  const valuesWithSymbols = biomarkers.filter(b => b.value && /[<>≤≥]/.test(b.value)).length;
  const totalValues = biomarkers.filter(b => b.value).length;
  
  if (totalValues === 0) return 0;
  
  return (valuesWithNumbers + valuesWithSymbols * 0.8) / totalValues;
}

function calculateUnitPrecisionScore(biomarkers) {
  if (biomarkers.length === 0) return 0;
  
  const withUnits = biomarkers.filter(b => b.unit && b.unit.trim()).length;
  const validUnits = biomarkers.filter(b => b.unit && isValidUnit(b.unit)).length;
  
  return withUnits > 0 ? validUnits / withUnits : 0;
}

function calculateNameStandardizationScore(biomarkers) {
  if (biomarkers.length === 0) return 0;
  
  const standardizedNames = biomarkers.filter(b => isStandardizedName(b.name)).length;
  return standardizedNames / biomarkers.length;
}

function isValidUnit(unit) {
  const validUnitPatterns = [
    /^mg\/d[Ll]$/,
    /^g\/d[Ll]$/,
    /^μg\/[Ll]$/,
    /^ng\/m[Ll]$/,
    /^pg\/m[Ll]$/,
    /^IU\/[Ll]$/,
    /^mIU\/[Ll]$/,
    /^U\/[Ll]$/,
    /^mmol\/[Ll]$/,
    /^μmol\/[Ll]$/,
    /^x10[³3]\/μ[Ll]$/,
    /^%$/,
    /^ratio$/i,
    /^index$/i
  ];
  
  return validUnitPatterns.some(pattern => pattern.test(unit.trim()));
}

function isStandardizedName(name) {
  // Check if the name follows medical naming conventions
  const standardPatterns = [
    /^[A-Z][a-z]+(\s[A-Z][a-z]+)*$/,  // Proper case
    /^[A-Z]{2,5}$/,  // Common abbreviations
    /^[A-Z][a-z]+\s[A-Z]{1,3}$/  // Name with abbreviation
  ];
  
  return standardPatterns.some(pattern => pattern.test(name.trim()));
}

function generatePrecisionRecommendations(extractedData, originalText) {
  const recommendations = [];
  const biomarkers = extractedData.biomarkers || [];
  
  // Text quality recommendations
  const ocrScore = assessOCRQuality(originalText);
  if (ocrScore < 0.8) {
    recommendations.push({
      type: "text_quality",
      priority: "high",
      message: "OCR quality is low. Consider using higher resolution scans or manual verification.",
      score: ocrScore
    });
  }
  
  // Unit completeness recommendations
  const missingUnits = biomarkers.filter(b => b.value && !b.unit).length;
  if (missingUnits > 0) {
    recommendations.push({
      type: "unit_completeness",
      priority: "medium",
      message: `${missingUnits} biomarkers are missing units. This may affect analysis precision.`,
      count: missingUnits
    });
  }
  
  // Low confidence extractions
  const lowConfidence = biomarkers.filter(b => b.confidence < 0.7).length;
  if (lowConfidence > 0) {
    recommendations.push({
      type: "extraction_confidence",
      priority: "medium",
      message: `${lowConfidence} biomarkers have low extraction confidence. Manual review recommended.`,
      count: lowConfidence
    });
  }
  
  // Missing numeric values
  const nonNumeric = biomarkers.filter(b => b.value && b.numericValue === null).length;
  if (nonNumeric > biomarkers.length * 0.3) {
    recommendations.push({
      type: "numeric_parsing",
      priority: "low",
      message: "Many values are non-numeric. This is normal for qualitative tests but verify if unexpected.",
      count: nonNumeric
    });
  }
  
  return recommendations;
}

function getCommonBiomarkerStats(dataset) {
  const markerCounts = {};
  dataset.forEach(entry => {
    const marker = entry.marker.toLowerCase();
    markerCounts[marker] = (markerCounts[marker] || 0) + 1;
  });
  
  return Object.entries(markerCounts)
    .sort(([,a], [,b]) => b - a)
    .slice(0, 20)
    .map(([marker, count]) => ({ marker, count }));
}

function extractUniqueUnits(dataset) {
  const units = new Set();
  
  dataset.forEach(entry => {
    // Extract units from min/max fields if they contain unit information
    const minStr = String(entry.min || '');
    const maxStr = String(entry.max || '');
    
    const unitPatterns = [
      /mg\/d[Ll]/g, /g\/d[Ll]/g, /μg\/[Ll]/g, /ng\/m[Ll]/g, /pg\/m[Ll]/g,
      /IU\/[Ll]/g, /mIU\/[Ll]/g, /U\/[Ll]/g, /mmol\/[Ll]/g, /μmol\/[Ll]/g,
      /x10[³3]\/μ[Ll]/g, /%/g, /ratio/gi, /index/gi
    ];
    
    unitPatterns.forEach(pattern => {
      const minMatches = minStr.match(pattern) || [];
      const maxMatches = maxStr.match(pattern) || [];
      [...minMatches, ...maxMatches].forEach(unit => units.add(unit));
    });
  });
  
  return Array.from(units).sort();
}

// ========== SERVER STARTUP ==========
app.listen(process.env.PORT || 5000, () => {
  console.log(`🎯 PRECISION-FOCUSED Medical Lab Server running on port ${process.env.PORT || 5000}`);
  console.log(`🧬 Enhanced for maximum biomarker extraction and analysis precision`);
  console.log(`🔬 Features: Ultra-precise extraction, advanced matching, validation metrics`);
  console.log(`📊 Using AI model: ${process.env.AI_MODEL || 'gemini-1.5-pro'} with precision optimizations`);
});