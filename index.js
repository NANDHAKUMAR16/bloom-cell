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
  dest: "uploads/", // Temporary directory for uploaded files
  limits: { fileSize: 25 * 1024 * 1024 }, // 25 MB file size limit
});

// Initialize Gemini AI client with API key from environment variables
const genAI = new GoogleGenerativeAI(process.env.API_KEY);

// Define the precise JSON schema for metadata and biomarkers extraction.
// This schema will be used to enforce valid JSON output from Gemini.
const EXTRACTION_RESPONSE_SCHEMA = {
  type: "object",
  properties: {
    metadata: {
      type: "object",
      properties: {
        patientName: { type: "string", nullable: true },
        age: { type: "string", nullable: true }, // Age can be "30 years", "unknown", etc.
        gender: { type: "string", nullable: true },
        dateOfBirth: { type: "string", nullable: true },
        reportGeneratedDate: { type: "string", nullable: true },
      },
      required: [], // No metadata fields are strictly required to be present in all reports
    },
    biomarkers: {
      type: "array",
      items: {
        type: "object",
        properties: {
          name: { type: "string" }, // Name of the biomarker (e.g., "Hemoglobin")
          group: { type: "string", nullable: true }, // Group it belongs to (e.g., "CBC")
          unit: { type: "string", nullable: true }, // Unit of measurement (e.g., "g/dL")
          value: { type: "string", nullable: true }, // Extracted value (e.g., "14.5", "Positive", "<10")
        },
        required: ["name"], // 'name' is the only strictly required field for a biomarker entry
      },
    },
  },
  required: ["metadata", "biomarkers"], // Both top-level 'metadata' and 'biomarkers' objects are required
};

// Prompt for initial data extraction.
// The prompt focuses on the task, as the schema handles the output format.
const STRUCTURED_EXTRACTION_PROMPT = `
You are a highly accurate clinical lab report extractor.
Your task is to extract patient details and a comprehensive list of all test results from the provided clinical lab report.

Extract all relevant information directly and precisely from the "SOURCE REPORT" text below.
For any field (in metadata or for a biomarker detail) that is missing, not applicable, or unclear in the report, set its value to null.
Ensure all extracted values are exact as they appear in the report, including units and any non-numeric results.

SOURCE REPORT:
<<<
{{RAW_TEXT}}
>>>
`.trim();

// ========== API Route for Extracting Biomarkers & Metadata ==========
app.post(
  "/extractReportGemini",
  upload.single("document"),
  async (req, res) => {
    // Validate file upload
    if (!req.file) {
      return res
        .status(400)
        .json({ success: false, error: "No file uploaded." });
    }
    const filePath = path.resolve(req.file.path);

    try {
      const buffer = fs.readFileSync(filePath);
      let rawText = "";

      // Determine file type and extract text content
      if (req.file.mimetype === "application/pdf") {
        const parsed = await pdfParse(buffer);
        rawText = parsed.text;
      } else if (
        req.file.mimetype ===
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document" || // .docx MIME type
        req.file.mimetype.includes("wordprocessingml") // Broader check for Word documents
      ) {
        const result = await mammoth.extractRawText({ path: filePath });
        rawText = result.value;
      } else {
        throw new Error(
          "Unsupported file type. Only PDF and DOCX documents are supported."
        );
      }

      // Clean up extracted text: replace multiple whitespaces with single space
      rawText = rawText.replace(/\s{3,}/g, " ").trim();

      // --- Input Text Length Check ---
      // Gemini 1.5 Pro has a 1 million token context window.
      // 1 token is approximately 4 characters. So, 1M tokens ≈ 4M characters.
      // We'll set a slightly lower practical limit to be safe.
      const MAX_CHARS_FOR_GEMINI_INPUT = 3_500_000; // ~875,000 tokens buffer
      if (rawText.length > MAX_CHARS_FOR_GEMINI_INPUT) {
        console.warn(
          `Input text length (${rawText.length} characters) exceeds recommended maximum for a single API call. ` +
            `Truncating to ${MAX_CHARS_FOR_GEMINI_INPUT} characters. For full accuracy on very long documents, consider chunking.`
        );
        rawText = rawText.substring(0, MAX_CHARS_FOR_GEMINI_INPUT);
      }

      // Get the Gemini model instance. Default to 'gemini-1.5-pro' if not specified in env.
      const model = genAI.getGenerativeModel({
        model: process.env.AI_MODEL || "gemini-1.5-pro",
      });

      // Prepare the prompt by injecting the raw text
      const prompt = STRUCTURED_EXTRACTION_PROMPT.replace(
        "{{RAW_TEXT}}",
        rawText
      );

      // Make the API call with structured output configuration
      const response = await model.generateContent({
        contents: [{ role: "user", parts: [{ text: prompt }] }],
        generationConfig: {
          temperature: 0.2, // Low temperature for more deterministic and consistent extraction
          maxOutputTokens: 65536, // *** INCREASED MAX OUTPUT TOKENS ***
          responseMimeType: "application/json", // Crucial: forces JSON output
          responseSchema: EXTRACTION_RESPONSE_SCHEMA, // Crucial: defines the exact JSON structure
        },
        safetySettings: [
          // Recommended safety settings for responsible AI
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

      // *** LOG RAW RESPONSE BEFORE PARSING FOR DEBUGGING ***
      const rawGeminiResponseText = response.response.text();

      // When responseMimeType is set, the response.response.text() is guaranteed to be valid JSON.
      // Direct parsing is safe here.
      const resultData = JSON.parse(rawGeminiResponseText);

      res.json({ success: true, data: resultData });
    } catch (err) {
      console.error("Gemini extraction failed:", err.message);
      // Log the raw Gemini error response for more detailed debugging
      if (err.response && err.response.text) {
        console.error(
          "Gemini raw error response (if available):",
          err.response.text()
        );
      }
      // Provide the raw text that caused JSON.parse error if it's the specific error
      if (
        err instanceof SyntaxError &&
        err.message.includes("Unexpected end of JSON input")
      ) {
        res.status(500).json({
          success: false,
          message: `Gemini extraction failed: ${err.message}. Raw response might be incomplete. Check server logs.`,
        });
      } else {
        res.status(500).json({
          success: false,
          message: `Gemini extraction failed: ${err.message}`,
        });
      }
    } finally {
      // Ensure the uploaded file is deleted from the temporary directory
      fs.unlink(filePath, (unlinkErr) => {
        if (unlinkErr)
          console.error("Error deleting uploaded file:", unlinkErr);
      });
    }
  }
);

// ========== Helper Functions for Dataset Loading & CSV Generation ==========

/**
 * Loads data from an XLSX file and converts it into an array of JSON objects.
 * @param {string} xlsxPath - The full path to the XLSX file.
 * @returns {Array<Object>} An array of objects, where each object represents a row.
 */
function loadDatasetRows(xlsxPath) {
  if (!fs.existsSync(xlsxPath)) {
    throw new Error(
      `Dataset file not found at: ${xlsxPath}. Please ensure 'All_Descriptions_Completed.xlsx' is in the same directory as your server script.`
    );
  }
  const wb = xlsx.readFile(xlsxPath);
  const sheet = wb.SheetNames[0]; // Assumes the first sheet contains the data
  return xlsx.utils.sheet_to_json(wb.Sheets[sheet], { defval: null });
}

/**
 * Generates a CSV string from an array of dataset rows.
 * @param {Array<Object>} rows - An array of objects representing the dataset rows.
 * @returns {string} A CSV formatted string.
 */
function generateDatasetCSV(rows) {
  // Define CSV headers, ensuring no unnecessary spaces
  const lines = ["Marker,Gender,Min,Max,Severity"];
  for (const r of rows) {
    // Standardize key access to lowercase for robustness against case variations in XLSX
    const marker = r["Blood Test Marker"] ?? r["blood test marker"] ?? "";
    const gender = r["Gender"] ?? r["gender"] ?? "";
    const min = r["Minimum"] ?? r["minimum"] ?? "";
    const max = r["Maximum"] ?? r["maximum"] ?? "";
    const severity =
      r["Severity Score (1 = mild, 5 = highly significant)"] ??
      r["severity score (1 = mild, 5 = highly significant)"] ??
      "";
    // Ensure no extra spaces in CSV lines
    lines.push(`${marker},${gender},${min},${max},${severity}`);
  }
  return lines.join("\n");
}

// Define the precise JSON schema for analysis output.
// This schema will be used to enforce valid JSON output from Gemini.
const ANALYSIS_RESPONSE_SCHEMA = {
  type: "object",
  properties: {
    evaluated: {
      type: "array",
      items: {
        type: "object",
        properties: {
          name: { type: "string" },
          value: { type: "string", nullable: true },
          unit: { type: "string", nullable: true },
          matched_marker: { type: "string", nullable: true },
          gender_used: { type: "string", nullable: true },
          min: { type: "string", nullable: true }, // Keep as string to handle "<10", ">100", "N/A"
          max: { type: "string", nullable: true }, // Keep as string
          severity: { type: "string", nullable: true }, // Keep as string for score or text
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
          value: { type: "string", nullable: true },
          unit: { type: "string", nullable: true },
          matched_marker: { type: "string", nullable: true }, // Should be null for unmatched
          gender_used: { type: "string", nullable: true }, // Should be null for unmatched
          min: { type: "string", nullable: true }, // Should be null for unmatched
          max: { type: "string", nullable: true }, // Should be null for unmatched
          severity: { type: "string", nullable: true }, // Should be null for unmatched
        },
        required: ["name"],
      },
    },
  },
  required: ["evaluated", "unmatched"], // Both top-level arrays are required
};

// ========== API Route for Analyzing Biomarkers with Dataset ==========
app.post("/analyzeWithDataset", async (req, res) => {
  try {
    const { biomarkers, metadata } = req.body.data;
    const gender = (metadata?.gender || "").toLowerCase();

    // Validate input biomarkers array
    if (!Array.isArray(biomarkers)) {
      return res.status(400).json({
        success: false,
        message: "Request body must contain a 'data.biomarkers' array.",
      });
    }

    // Load and prepare the dataset
    const datasetPath = path.join(__dirname, "All_Descriptions_Completed.xlsx");
    const dataset = loadDatasetRows(datasetPath);
    const datasetCSV = generateDatasetCSV(dataset);

    // Get the Gemini model instance. Default to 'gemini-1.5-pro' if not specified in env.
    const model = genAI.getGenerativeModel({
      model: process.env.AI_MODEL || "gemini-1.5-pro",
    });
    console.log(model);
    const biomarkerText = JSON.stringify(biomarkers, null, 2); // Pretty print for better context for the LLM

    // Prompt for analysis, simplified for structured output
    const prompt = `
      You are a clinical lab analysis assistant.
      Your primary task is to match each patient biomarker to the most appropriate reference row in the provided dataset and determine if its value is within the defined normal range.

      ### Matching Logic:
      - Use medical synonym awareness, abbreviation expansion, and fuzzy name similarity (e.g., "WBC" should match "White Blood Cells").
      - Prioritize dataset rows where the 'Gender' field precisely matches the patient’s gender.
      - If no exact gender match (e.g., "male" or "female") exists, then use the row where 'Gender' is explicitly "both".
      - Absolutely never match a biomarker using a dataset row with a mismatched gender (e.g., do not use a "female" range for a "male" patient).

      ### Unit Precision Rules:
      - Only match a biomarker if its 'unit' is identical to the dataset's 'unit' (e.g., "ng/mL" must match "ng/mL").
      - If units differ but are clearly convertible (e.g., "ng/mL" vs "μg/mL"), you may match only if clinically appropriate and you are highly confident in the conversion. State the converted unit in 'unit' field.
      - Never match biomarkers where the units are not compatible (e.g., "pg/mL" vs "mmol/L").

      Include a biomarker in the "evaluated" list ONLY if you can successfully match it to a dataset entry and populate at least one of these fields: matched_marker, min, max, or severity. All other biomarkers must be placed in the "unmatched" list. For unmatched items, ensure 'matched_marker', 'gender_used', 'min', 'max', 'severity' are explicitly set to null.

      === PATIENT GENDER ===
      ${gender}

      === PATIENT BIOMARKERS (JSON) ===
      ${biomarkerText}

      === REFERENCE DATASET (CSV) ===
      ${datasetCSV}
      `.trim();

    const response = await model.generateContent({
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: {
        temperature: 0.2, // Low temperature for deterministic analysis
        maxOutputTokens: 65536, // *** INCREASED MAX OUTPUT TOKENS ***
        responseMimeType: "application/json", // Crucial: forces JSON output
        responseSchema: ANALYSIS_RESPONSE_SCHEMA, // Crucial: defines the exact JSON structure
      },
      safetySettings: [
        // Recommended safety settings for responsible AI
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

    // *** LOG RAW RESPONSE BEFORE PARSING FOR DEBUGGING ***
    const rawGeminiResponseText = response.response.text();

    // Direct JSON parsing is safe due to responseMimeType and responseSchema
    const modelOutput = JSON.parse(rawGeminiResponseText);

    // --- Post-processing for 'evaluated' vs 'unmatched' ---
    // This logic strictly enforces your business rules, overriding potential LLM deviations.
    const evaluated = [];
    const unmatched = [];

    // Combine all entries from model's output for re-evaluation
    const allEntries = [
      ...(Array.isArray(modelOutput.evaluated) ? modelOutput.evaluated : []),
      ...(Array.isArray(modelOutput.unmatched) ? modelOutput.unmatched : []),
    ];

    allEntries.forEach((entry) => {
      // Determine if it's a valid match based on your criteria
      const isValidMatch =
        entry.matched_marker !== null &&
        entry.matched_marker !== "" && // Must have a matched marker
        (entry.min !== null || entry.max !== null || entry.severity !== null); // And at least one range/severity detail

      if (isValidMatch) {
        let isWithinRange = null; // Default to null if range cannot be determined numerically

        const patientValue = parseFloat(entry.value);
        const minRange = parseFloat(entry.min);
        const maxRange = parseFloat(entry.max);

        if (!isNaN(patientValue)) {
          // Only proceed if patient value is a number
          if (!isNaN(minRange) && !isNaN(maxRange)) {
            isWithinRange =
              patientValue >= minRange && patientValue <= maxRange;
          } else if (!isNaN(minRange) && isNaN(maxRange)) {
            isWithinRange = patientValue >= minRange; // Only min provided
          } else if (isNaN(minRange) && !isNaN(maxRange)) {
            isWithinRange = patientValue <= maxRange; // Only max provided
          } else {
            // If value is number but range min/max are not numbers (e.g., "Positive"),
            // For now, if numerical ranges are missing, and it was a valid match, assume true.
            isWithinRange = true;
          }
        }

        evaluated.push({
          ...entry,
        });
      } else {
        // If not a valid match, ensure it's in the unmatched list with nullified fields
        unmatched.push({
          name: entry.name || null, // Ensure name is present, or null
          value: entry.value || null,
          unit: entry.unit || null,
          matched_marker: null,
          gender_used: null,
          min: null,
          max: null,
          severity: null,
        });
      }
    });

    res.json({ success: true, evaluated, unmatched });
  } catch (err) {
    console.error("Dataset analysis failed:", err.message);
    // Log the raw Gemini error response for more detailed debugging
    if (err.response && err.response.text) {
      console.error(
        "Gemini raw error response (if available):",
        err.response.text()
      );
    }
    // Provide the raw text that caused JSON.parse error if it's the specific error
    if (
      err instanceof SyntaxError &&
      err.message.includes("Unexpected end of JSON input")
    ) {
      res.status(500).json({
        success: false,
        message: `Dataset analysis failed: ${err.message}. Raw response might be incomplete. Check server logs.`,
      });
    } else {
      res.status(500).json({
        success: false,
        message: `Dataset analysis failed: ${err.message}`,
      });
    }
  }
});

// ========== Server ==========
app.listen(process.env.PORT || 5000, () => {
  console.log(`Server running on port ${process.env.PORT || 5000}`);
});
