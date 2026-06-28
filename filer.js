import express from "express";
import multer from "multer";
import XLSX from "xlsx";
import cors from "cors";
import fs from "fs";
import { GoogleGenerativeAI } from "@google/generative-ai";

const app = express();
const upload = multer({ dest: "uploads/" });

app.use(cors());

const genAI = new GoogleGenerativeAI(process.env.AIzaSyD2UAKrhkrWtR4VYnQ5UJp7rVQHtEsZBqc);

app.post("/convert", upload.single("image"), async (req, res) => {
  try {
    const imagePath = req.file.path;

    const imageBuffer = fs.readFileSync(imagePath);
    const base64Image = imageBuffer.toString("base64");

    const model = genAI.getGenerativeModel({
      model: "gemini-1.5-flash", // Fast model with generous free-tier limits
    });

    // Send the image to Gemini and request structured JSON output
    const result = await model.generateContent([
      {
        text: `
Extract all tabular data from this image.

Return ONLY valid JSON.
Format:
- Array of objects
- Each object = one row
- Keys = column headers

Rules:
- Clean headers
- Fill missing values if obvious
- Keep structure consistent

NO explanation.
ONLY JSON.
        `,
      },
      {
        inlineData: {
          mimeType: "image/png",
          data: base64Image,
        },
      },
    ]);

    let text = result.response.text();

    // Strip markdown code fences from the model response
    text = text
      .replace(/```json/g, "")
      .replace(/```/g, "")
      .trim();

    const data = JSON.parse(text);

    // Convert the parsed JSON rows into an Excel workbook
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    const buffer = XLSX.write(workbook, {
      type: "buffer",
      bookType: "xlsx",
    });

    fs.unlinkSync(imagePath);

    res.setHeader(
      "Content-Disposition",
      "attachment; filename=output.xlsx"
    );
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).send("Error processing image");
  }
});

app.listen(5000, () => {
  console.log("Server running on port 5000");
});