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

// ========== SIMPLIFIED EXTRACTION CONFIGURATION ==========

// Simplified JSON schema for biomarker extraction
const SIMPLIFIED_EXTRACTION_SCHEMA = {
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
          enum: [null, "male", "female", "both", "unknown"],
        },
        dateOfBirth: { type: "string", nullable: true },
        reportGeneratedDate: { type: "string", nullable: true },
        labName: { type: "string", nullable: true },
        doctorName: { type: "string", nullable: true },
        reportId: { type: "string", nullable: true },
      },
      required: [],
    },
    biomarkers: {
      type: "array",
      items: {
        type: "object",
        properties: {
          name: {
            type: "string",
            description:
              "Exact name of the biomarker as it appears in the report",
          },
          group: {
            type: "string",
            nullable: true,
            description:
              "Test panel or group this biomarker belongs to (e.g., CBC, CMP, Lipid Panel)",
          },
          unit: {
            type: "string",
            nullable: true,
            description: "Exact unit of measurement as written in the report",
          },
          value: {
            type: "string",
            nullable: true,
            description:
              "Exact value as it appears, including any symbols like <, >, or ranges",
          },
          numericValue: {
            type: "number",
            nullable: true,
            description:
              "Parsed numeric value if the value can be converted to a number",
          },
          referenceRange: {
            type: "string",
            nullable: true,
            description:
              "Reference range as stated in the report, if available",
          },
        },
        required: ["name"],
      },
    },
  },
  required: ["metadata", "biomarkers"],
};

// Simplified extraction prompt
const SIMPLIFIED_EXTRACTION_PROMPT = `
You are an expert medical laboratory technologist. Extract biomarker data with maximum precision.

EXTRACTION REQUIREMENTS:

1. **EXACT VALUE EXTRACTION**:
   - Copy values EXACTLY as they appear: "14.5", "<0.5", ">100", "1.2-3.4", "NONE SEEN"
   - Never round, estimate, or modify values
   - Include all symbols: <, >, ±, ~, etc.

2. **UNIT PRECISION**:
   - Copy units EXACTLY: "mg/dL", "μg/L", "ng/mL", "IU/L", "mmol/L"
   - Don't standardize or convert units
   - Include all special characters: μ, ², ³, /, etc.

3. **NAME STANDARDIZATION**:
   - Use the most complete name available in the report
   - Common standardizations:
     * "Hemoglobin" (not "Hgb", "Hb", "HEMOGLOBIN")
     * "White Blood Cell Count" (not "WBC", "Leukocytes")
     * "Total Cholesterol" (not "CHOL", "Cholesterol")
     * "Glucose" (not "GLU", "Blood Sugar")

4. **COMPREHENSIVE EXTRACTION**:
   Extract ALL test results including:
   - Complete Blood Count (CBC) with differential
   - Basic/Comprehensive Metabolic Panel (BMP/CMP)
   - Lipid profiles, Liver function tests, Kidney function tests
   - Cardiac markers, Inflammatory markers, Thyroid function
   - Diabetes markers, Tumor markers, Coagulation studies
   - Urinalysis components, Immunology/Serology results
   - Any other laboratory values

SOURCE REPORT:
<<<
{{RAW_TEXT}}
>>>
`.trim();

// Enhanced analysis schema (simplified)
const SIMPLIFIED_ANALYSIS_SCHEMA = {
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
          gender_used: { type: "string", nullable: true },
          min_reference: { type: "string", nullable: true },
          max_reference: { type: "string", nullable: true },
          min_numeric: { type: "number", nullable: true },
          max_numeric: { type: "number", nullable: true },
          severity_score: { type: "string", nullable: true },
          isWithinRange: { type: "boolean", nullable: true },
          clinical_significance: {
            type: "string",
            nullable: true,
            enum: [null, "normal", "borderline", "abnormal", "critical"],
          },
        },
        required: ["name"],
      },
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
            ],
          },
        },
        required: ["name", "unmatch_reason"],
      },
    },
  },
  required: ["evaluated", "unmatched"],
};

// Simplified analysis prompt
const SIMPLIFIED_ANALYSIS_PROMPT = `
You are a clinical laboratory data analyst. Match biomarkers with reference ranges accurately.

MATCHING PROTOCOL:

1. **BIOMARKER NAME MATCHING**:
   - Use medical knowledge for synonym matching
   - Common matches to recognize:
     * Hemoglobin = Hgb = Hb = HGB
     * White Blood Cell Count = WBC = Leukocytes = White Blood Cells
     * Total Cholesterol = Cholesterol = CHOL = TC
     * Alanine Aminotransferase = ALT = SGPT
     * Aspartate Aminotransferase = AST = SGOT
     * Blood Urea Nitrogen = BUN = Urea Nitrogen
     * Creatinine = Cr = CREAT
     * Glucose = GLU = Blood Sugar = BG

2. **GENDER MATCHING HIERARCHY**:
   - Priority 1: Exact gender match ("male" for male patient)
   - Priority 2: "both" gender entry
   - Priority 3: No match (mark as unmatched)

3. **CLINICAL SIGNIFICANCE ASSESSMENT**:
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
`.trim();

// ========== HELPER FUNCTIONS ==========

/**
 * Enhanced text preprocessing for better OCR/extraction accuracy
 */
function preprocessTextForPrecision(rawText) {
  let processed = rawText.replace(/[ \t]{2,}/g, " ");

  processed = processed
    .replace(/[Oo](?=\d)/g, "0") // O -> 0 before digits
    .replace(/[Il](?=\d)/g, "1") // I,l -> 1 before digits
    .replace(/(\d)[Oo](\d)/g, "$10$2") // 0 between digits
    .replace(/(\d)[Il](\d)/g, "$11$2") // 1 between digits
    .replace(/mg\/d[Il]/g, "mg/dL") // Common mg/dL OCR error
    .replace(/μg\/[Il]/g, "μg/L") // Common μg/L OCR error
    .replace(/mcg\/[Il]/g, "mcg/L") // Common mcg/L OCR error
    .replace(/([A-Za-z]+)\s*[:]\s*([<>]?\s*\d+\.?\d*)/g, "$1: $2")
    .replace(/\s*:\s*/g, ": ")
    .replace(/([<>])\s+(\d)/g, "$1$2") // Remove space after < or >
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  return processed;
}

/**
 * Advanced biomarker name standardization
 */
function standardizeBiomarkerName(name) {
  const standardizations = {
    // Hematology
    wbc: "White Blood Cell Count",
    "white blood cells": "White Blood Cell Count",
    leukocytes: "White Blood Cell Count",
    rbc: "Red Blood Cell Count",
    "red blood cells": "Red Blood Cell Count",
    hgb: "Hemoglobin",
    hb: "Hemoglobin",
    hct: "Hematocrit",
    plt: "Platelet Count",
    platelets: "Platelet Count",
    mcv: "Mean Corpuscular Volume",
    mch: "Mean Corpuscular Hemoglobin",
    mchc: "Mean Corpuscular Hemoglobin Concentration",
    rdw: "Red Cell Distribution Width",

    // Chemistry
    glu: "Glucose",
    glucose: "Glucose",
    "blood sugar": "Glucose",
    bun: "Blood Urea Nitrogen",
    cr: "Creatinine",
    creat: "Creatinine",
    na: "Sodium",
    k: "Potassium",
    cl: "Chloride",
    co2: "Carbon Dioxide",
    chol: "Total Cholesterol",
    cholesterol: "Total Cholesterol",
    hdl: "HDL Cholesterol",
    ldl: "LDL Cholesterol",
    trig: "Triglycerides",
    triglycerides: "Triglycerides",

    // Liver function
    alt: "Alanine Aminotransferase",
    sgpt: "Alanine Aminotransferase",
    ast: "Aspartate Aminotransferase",
    sgot: "Aspartate Aminotransferase",
    alkp: "Alkaline Phosphatase",
    alp: "Alkaline Phosphatase",
    tbili: "Total Bilirubin",
    bilirubin: "Total Bilirubin",

    // Thyroid
    tsh: "Thyroid Stimulating Hormone",
    t4: "Thyroxine",
    t3: "Triiodothyronine",

    // Cardiac
    ck: "Creatine Kinase",
    cpk: "Creatine Kinase",
    "ck-mb": "Creatine Kinase-MB",
    "troponin i": "Troponin I",
    "troponin t": "Troponin T",

    // Inflammatory
    crp: "C-Reactive Protein",
    esr: "Erythrocyte Sedimentation Rate",

    // Diabetes
    hba1c: "Hemoglobin A1c",
    a1c: "Hemoglobin A1c",
  };

  const lowerName = name.toLowerCase().trim();
  return standardizations[lowerName] || name;
}

/**
 * Extract unit from reference range
 */
function extractUnitFromReferenceRange(referenceRange) {
  if (!referenceRange || typeof referenceRange !== "string") {
    return null;
  }

  // Common unit patterns in reference ranges
  const unitPatterns = [
    /mg\/d[Ll]/g,
    /g\/d[Ll]/g,
    /μg\/[Ll]/g,
    /mcg\/[Ll]/g,
    /ng\/m[Ll]/g,
    /pg\/m[Ll]/g,
    /IU\/[Ll]/g,
    /mIU\/[Ll]/g,
    /U\/[Ll]/g,
    /mmol\/[Ll]/g,
    /μmol\/[Ll]/g,
    /nmol\/[Ll]/g,
    /x10[³3]\/μ[Ll]/g,
    /K\/μ[Ll]/g,
    /10[³3]\/μ[Ll]/g,
    /cells\/μ[Ll]/g,
    /\/μ[Ll]/g,
    /mg\/24h/g,
    /g\/24h/g,
    /mmHg/g,
    /bpm/g,
    /°[CF]/g,
    /%/g,
    /ratio/gi,
    /index/gi,
    /sec/g,
    /min/g,
    /fl/gi,
    /pg/gi,
  ];

  // Try to find unit patterns in the reference range
  for (const pattern of unitPatterns) {
    const matches = referenceRange.match(pattern);
    if (matches && matches.length > 0) {
      return matches[0]; // Return the first match
    }
  }

  // Fallback: look for any pattern that looks like a unit
  const fallbackPattern =
    /\b[a-zA-Z]+\/[a-zA-Z]+\b|\b[μ]?[a-zA-Z]+\/[a-zA-Z]+\b|\b[a-zA-Z]{1,5}\b(?=\s|$)/g;
  const fallbackMatches = referenceRange.match(fallbackPattern);

  if (fallbackMatches) {
    // Filter out common non-unit words
    const nonUnits = [
      "normal",
      "abnormal",
      "high",
      "low",
      "range",
      "reference",
      "adult",
      "male",
      "female",
      "years",
      "old",
      "seen",
      "none",
      "few",
      "many",
      "rare",
      "moderate",
      "negative",
      "positive",
    ];
    const filteredMatches = fallbackMatches.filter(
      (match) =>
        !nonUnits.includes(match.toLowerCase()) &&
        match.length <= 8 &&
        /[a-zA-Z]/.test(match) // Contains at least one letter
    );

    if (filteredMatches.length > 0) {
      return filteredMatches[0];
    }
  }

  return null;
}

/**
 * Enhanced post-processing function for biomarkers
 */
function postProcessBiomarkers(biomarkers) {
  return biomarkers.map((biomarker) => {
    const processedBiomarker = {
      name: standardizeBiomarkerName(biomarker.name || ""),
      group: biomarker.group || null,
      unit: biomarker.unit || null,
      value: biomarker.value || null,
      numericValue: null,
      referenceRange: biomarker.referenceRange || null,
    };

    // Extract unit from reference range if unit is null
    if (!processedBiomarker.unit && processedBiomarker.referenceRange) {
      const extractedUnit = extractUnitFromReferenceRange(
        processedBiomarker.referenceRange
      );
      if (extractedUnit) {
        processedBiomarker.unit = extractedUnit;
      }
    }

    // Parse numeric values with precision
    if (biomarker.value) {
      const cleanValue = biomarker.value.replace(/[<>≤≥]/g, "").trim();
      const numericMatch = cleanValue.match(/^(\d+\.?\d*)/);
      if (numericMatch && !isNaN(parseFloat(numericMatch[1]))) {
        processedBiomarker.numericValue = parseFloat(numericMatch[1]);
      }
    }

    return processedBiomarker;
  });
}

/**
 * Load dataset with enhanced error handling
 */
function loadDatasetWithPrecision(xlsxPath) {
  if (!fs.existsSync(xlsxPath)) {
    throw new Error(`Dataset file not found: ${xlsxPath}`);
  }

  const wb = xlsx.readFile(xlsxPath);
  const sheetName = wb.SheetNames[0];
  const worksheet = wb.Sheets[sheetName];

  const rawData = xlsx.utils.sheet_to_json(worksheet, {
    defval: null,
    raw: false,
    dateNF: "yyyy-mm-dd",
  });

  const cleanedData = rawData
    .filter((row) => {
      const marker = row["Blood Test Marker"] || row["blood test marker"];
      return marker && marker.trim().length > 0;
    })
    .map((row) => ({
      marker: (
        row["Blood Test Marker"] ||
        row["blood test marker"] ||
        ""
      ).trim(),
      gender: (row["Gender"] || row["gender"] || "").toLowerCase().trim(),
      min: row["Minimum"] || row["minimum"] || "",
      max: row["Maximum"] || row["maximum"] || "",
      severity:
        row["Severity Score (1 = mild, 5 = highly significant)"] ||
        row["severity score (1 = mild, 5 = highly significant)"] ||
        "",
    }));

  console.log(`Loaded ${cleanedData.length} validated dataset entries`);
  return cleanedData;
}

/**
 * Generate CSV for analysis
 */
function generatePrecisionCSV(datasetRows) {
  const headers = [
    "Marker",
    "Gender",
    "Min",
    "Max",
    "Group",
    "High",
    "Low",
    "Severity",
  ];
  const csvLines = [headers.join(",")];

  datasetRows.forEach((row) => {
    const csvRow = [
      `"${row.marker.replace(/"/g, '""')}"`,
      `"${row.gender}"`,
      `"${row.min}"`,
      `"${row.max}"`,
      `"${row.group}"`,
      `"${row.high}"`,
      `"${row.low}"`,
      `"${row.severity}"`,
    ];
    csvLines.push(csvRow.join(","));
  });

  return csvLines.join("\n");
}

// ========== API ROUTES ==========

app.post(
  "/extractReportGemini",
  upload.single("document"),
  async (req, res) => {
    if (!req.file) {
      return res
        .status(400)
        .json({ success: false, error: "No file uploaded." });
    }

    const filePath = path.resolve(req.file.path);

    try {
      const buffer = fs.readFileSync(filePath);
      let rawText = "";

      // Extract text based on file type
      if (req.file.mimetype === "application/pdf") {
        const parsed = await pdfParse(buffer);
        rawText = parsed.text;
      } else if (req.file.mimetype.includes("wordprocessingml")) {
        const result = await mammoth.extractRawText({ path: filePath });
        rawText = result.value;
      } else {
        throw new Error("Unsupported file type. Only PDF and DOCX supported.");
      }

      // Apply preprocessing
      rawText = preprocessTextForPrecision(rawText);

      // Check text quality
      if (rawText.length < 100) {
        return res.status(400).json({
          success: false,
          error: "Insufficient readable text extracted from document.",
        });
      }

      // Truncate if too long
      const MAX_CHARS = 3_500_000;
      if (rawText.length > MAX_CHARS) {
        console.warn(
          `Text truncated from ${rawText.length} to ${MAX_CHARS} characters`
        );
        rawText = rawText.substring(0, MAX_CHARS);
      }

      // Get model and generate response
      const model = genAI.getGenerativeModel({
        model: process.env.AI_MODEL || "gemini-1.5-pro",
      });

      const prompt = SIMPLIFIED_EXTRACTION_PROMPT.replace(
        "{{RAW_TEXT}}",
        rawText
      );

      const response = await model.generateContent({
        contents: [{ role: "user", parts: [{ text: prompt }] }],
        generationConfig: {
          temperature: 0.1,
          maxOutputTokens: 65536,
          responseMimeType: "application/json",
          responseSchema: SIMPLIFIED_EXTRACTION_SCHEMA,
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

      // Post-process biomarkers with unit extraction from reference ranges
      if (resultData.biomarkers) {
        resultData.biomarkers = postProcessBiomarkers(resultData.biomarkers);
      }

      // Add total count of biomarkers
      const totalBiomarkers = resultData.biomarkers
        ? resultData.biomarkers.length
        : 0;

      res.json({
        success: true,
        data: {
          ...resultData,
          totalBiomarkers: totalBiomarkers,
        },
      });
    } catch (err) {
      console.error("Extraction failed:", err.message);
      res.status(500).json({
        success: false,
        message: `Extraction failed: ${err.message}`,
        debug: process.env.NODE_ENV === "development" ? err.stack : undefined,
      });
    } finally {
      // Clean up file
      fs.unlink(filePath, (unlinkErr) => {
        if (unlinkErr) console.error("Error deleting file:", unlinkErr);
      });
    }
  }
);

app.post("/analyzeWithDataset", async (req, res) => {
  try {
    const { biomarkers, metadata } = req.body.data;

    if (!Array.isArray(biomarkers)) {
      return res.status(400).json({
        success: false,
        message: "Request body must contain a 'data.biomarkers' array.",
      });
    }

    const gender = (metadata?.gender || "").toLowerCase();

    // Load dataset
    const datasetPath = path.join(__dirname, "All_Descriptions_Completed.xlsx");
    const dataset = loadDatasetWithPrecision(datasetPath);
    const datasetCSV = generatePrecisionCSV(dataset);

    // Process biomarkers
    const processedBiomarkers = biomarkers.map((biomarker) => ({
      ...biomarker,
      name: standardizeBiomarkerName(biomarker.name || ""),
      numericValue:
        biomarker.value &&
        !isNaN(parseFloat(biomarker.value.replace(/[<>]/g, "")))
          ? parseFloat(biomarker.value.replace(/[<>]/g, ""))
          : null,
    }));

    const model = genAI.getGenerativeModel({
      model: process.env.AI_MODEL || "gemini-1.5-pro",
    });

    const prompt = SIMPLIFIED_ANALYSIS_PROMPT.replace("{{GENDER}}", gender)
      .replace("{{BIOMARKERS}}", JSON.stringify(processedBiomarkers, null, 2))
      .replace("{{DATASET_CSV}}", datasetCSV);

    const response = await model.generateContent({
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: {
        temperature: 0.1,
        maxOutputTokens: 65536,
        responseMimeType: "application/json",
        responseSchema: SIMPLIFIED_ANALYSIS_SCHEMA,
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

    res.json({ success: true, ...analysisResult });
  } catch (err) {
    console.error("Analysis failed:", err.message);

    if (err.message.includes("Unexpected end of JSON")) {
      res.status(500).json({
        success: false,
        message: "AI analysis was incomplete. Please try again.",
        error_type: "incomplete_json_response",
      });
    } else if (err.code === "ENOENT") {
      res.status(500).json({
        success: false,
        message:
          "Dataset file not found. Please ensure 'All_Descriptions_Completed.xlsx' is available.",
        error_type: "missing_dataset",
      });
    } else {
      res.status(500).json({
        success: false,
        message: `Analysis failed: ${err.message}`,
        error_type: "analysis_error",
        debug: process.env.NODE_ENV === "development" ? err.stack : undefined,
      });
    }
  }
});

// Health check endpoint
app.get("/health", (req, res) => {
  res.json({
    success: true,
    message: "Server is running",
    timestamp: new Date().toISOString(),
  });
});

app.use("/", (req, res) => {
  console.log("Hello");
  res.json({
    message: "Welcome to the Medical Lab Server",
  });
});
// ========== SERVER STARTUP ==========
app.listen(process.env.PORT || 5000, () => {
  console.log(`Medical Lab Server running on port ${process.env.PORT || 5000}`);
  console.log(`Enhanced for simplified biomarker extraction and analysis`);
  console.log(`Using AI model: ${process.env.AI_MODEL || "gemini-1.5-pro"}`);
});
