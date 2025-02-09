import express from "express";
import multer from "multer";
import xlsx from "xlsx";
import pkg from "pg";
import cors from "cors";

const { Pool } = pkg;

const app = express();
const PORT = 3000;

// Enhanced CORS setup
const corsOptions = {
  origin: "http://localhost:4200", // Replace with your frontend URL
  methods: ["GET", "POST", "PUT", "DELETE"],
  allowedHeaders: ["Content-Type", "Authorization"],
  credentials: true,
};

app.use(cors(corsOptions));
app.options("*", cors(corsOptions)); // Handle preflight requests

// PostgreSQL connection settings
const pool = new Pool({
  user: "postgres",
  host: "localhost",
  database: "users",
  password: "Animesh@123",
  port: 5432,
});

pool.connect((err, client, release) => {
  if (err) {
    return console.error("Error acquiring client", err.stack);
  }
  console.log("Connected to PostgreSQL!");
  release();
});

// Multer setup for file uploads
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Route to upload Excel file

app.post("/upload", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).send("No file uploaded.");
  }

  try {
    const workbook = xlsx.read(req.file.buffer, { type: "buffer" });
    const allData = [];

    // Loop through all sheets
    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      const sheetData = xlsx.utils.sheet_to_json(worksheet);

      allData.push({ sheetName, sheetData });
    }

    const client = await pool.connect();
    const insertedData = [];
    const duplicateEmails = [];

    try {
      // Process each sheet's data
      for (const sheet of allData) {
        for (const row of sheet.sheetData) {
          const { Name, Email, Age, isRedFlag } = row;

          // Validate required fields
          if (!Name || !Email || Age === undefined || isRedFlag === undefined) {
            console.warn(
              `Skipping row due to missing data: ${JSON.stringify(row)}`
            );
            continue;
          }

          const result = await client.query(
            `INSERT INTO users (name, email, age, isRedFlag) 
             VALUES ($1, $2, $3, $4) 
             ON CONFLICT (email) DO NOTHING 
             RETURNING *`,
            [Name, Email, Age, isRedFlag]
          );

          if (result.rows.length > 0) {
            // insertedData.push(result.rows[0]);

            insertedData.push({
              sheetName: sheet.sheetName, // Add the sheetName to the inserted data
              data: result.rows[0],
            });

            // console.log("Inserted data: ", sheet.sheetName);
          } else {
            // duplicateEmails.push(Email);
            duplicateEmails.push({
              sheetName: sheet.sheetName, // Add the sheetName to the duplicate emails list
              email: Email,
            });
          }
        }
      }

      res.setHeader("Content-Type", "application/json");
      res.status(201).json({
        message: "File processed and data inserted successfully.",
        status: "success",
        data: insertedData,
        serverMessage:
          duplicateEmails.length > 0
            ? `The following emails were duplicates and not inserted: ${duplicateEmails.join(
                ", "
              )}`
            : "", // If there are no duplicates, an empty string is returned
      });
    } finally {
      client.release(); // Ensure the client is always released
    }
  } catch (error) {
    console.error("Error processing file", error);
    res.status(500).json({
      message: "Error processing file.",
      status: "error",
      serverMessage: error.message,
    });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
