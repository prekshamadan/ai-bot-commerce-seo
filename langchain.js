import { ChatOpenAI } from "@langchain/openai";
import { ChatPromptTemplate } from "@langchain/core/prompts";

import fs from 'fs';
import csv from 'csv-parser';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';

import { promisify } from 'util';
import path from 'path';

// Import dotenv and load the environment variables
import { config } from 'dotenv';
import { Console } from "console";
config();

// Ideally, store sensitive information like API keys in environment variables
const openAIApiKey = process.env.OPENAI_API_KEY;

// Initialize chat model with the OpenAI API key
const chatModel = new ChatOpenAI({
    openAIApiKey: openAIApiKey,
   // modelName: 'gpt-4o',
    maxConcurrency: 5
});

// Promisify fs.readFile and fs.writeFile for use with async/await
const readFileAsync = promisify(fs.readFile);
const writeFileAsync = promisify(fs.writeFile);

// Function to read CSV file asynchronously and return product data
async function readCsvFile(filePath) {
    const products = [];
    return new Promise((resolve, reject) => {
        fs.createReadStream(filePath)
            .pipe(csv())
            .on('data', (row) => products.push(row))
            .on('end', () => resolve(products))
            .on('error', (err) => reject(err));
    });
}

//Generate AI content for a single product
const generateResponse = async (name, url) => {
    //input
    const product_name = name;
    const product_url = url;

    //fetch page description to be fed to the main prompt
    const descriptionPrompt = ChatPromptTemplate.fromMessages([
        ["system", `Extract product description (if any) from a product url
        - Don't add any reference links in your answer
        - Try not to change any language and give answer in a single paragraph`],
        ["user", "{input}"],
        ]);
    const descriptionChain = descriptionPrompt.pipe(chatModel);
    const product_description = await descriptionChain.invoke({ input: `Product url is ${product_url}`});

    //console.log(product_description.lc_kwargs.content);

    console.log(`description generated for ${name}`);

    //main prompt to generate content
    const mainPrompt = ChatPromptTemplate.fromMessages([
        ["system", `You are a content and SEO writer who is suppose to write product description for an ecommerce website selling e-commerce products. Keep the below rules as do's and dont's while writing the descriptions.
                
        Do’s for Writing Product Descriptions:

            1.	Provide Accurate Information: Conduct thorough research on the product to present accurate details about its features, specifications, and benefits.
            2.	Use Clear and Understandable Language: Ensure the language is straightforward, making the product’s features and advantages easy to understand for potential buyers.
            3.	Focus on Practical Applications: Emphasize how the product can be used in everyday life, showcasing its real-world relevance and value.
            4.	Highlight Key Selling Points: Clearly outline the unique aspects and advantages of the product compared to alternatives.
            5.	Be Transparent About Variability: Explain that performance, functionality, or results may vary based on individual usage or environmental factors.

        Don’ts for Writing Product Descriptions:

            1.	Avoid Overhyping: Refrain from exaggerated claims or overpromising the product’s features, benefits, or results.
            2.	Avoid Misleading Language: Do not use ambiguous or potentially misleading terms that could confuse the customer.
            3.	No False Claims: Avoid making unverifiable claims about endorsements, exclusivity, or guaranteed success.
            4.	Don’t Use Vague Jargon: Keep descriptions grounded, avoiding overly technical or vague language unless essential.
            5.	Avoid Omitting Disclaimers: Include necessary warnings, disclaimers, or information regarding compatibility or potential limitations.
            6.	Avoid Keyword Stuffing: Incorporate SEO keywords naturally without stuffing or disrupting the flow of the description.

        Avoid These Terms in Your Writing:

            •	Guaranteed satisfaction
            •	Perfect choice
            •	Best in the market
            •	Once-in-a-lifetime
            •	Must-have
            •	Secret design
            •	100% effective
            •	Exclusive
            •	No downsides
            •	Universal appeal
            •	Limited only to you
            •	All-in-one solution
            •	Foolproof system
            •	Perfect results
                
                You will be given a product and in some cases the existing content on the website. You need to write a detailed product description about the product in the following format:
                Heading of the page with product name
                Product Description (2-3 paragraphs with at least 70 words in each paragraph)
                Frequently Asked Questions: Add five to Six commonly asked questions and their answers (10-20 words for each answer)
                Key benefits: Summarize key benefits in 4 pointers (30-50 words)
                Direction for use (30-50 words)
                Safety information
                Other information (30-50 words)
                Meta title including complete product name and other keywords (50-70 characters)
                Meta Description including main SEO keywords (130- 170 characters)`
        ],
        ["user", "{input}"],
        ]);
    const mainChain = mainPrompt.pipe(chatModel);

    const response = await mainChain.invoke({
            input: `Product name is ${product_name}
            Existing description is ${product_description}
            `
        });

    //console.log(response.lc_kwargs.content);
    console.log(`content generated for ${name}`);
    return (response.lc_kwargs.content);
}

// Function to create and save the Word document
async function createWordDocument(products) {
    try {
        const templateContent = await readFileAsync('template.docx');
        const zip = new PizZip(templateContent);
        const doc = new Docxtemplater(zip, { 
            paragraphLoop: true, 
            linebreaks: true
        });

        //console.log({products});
        doc.setData({ products });
        doc.render();

        const dateTimeString = getFormattedDateTime();
        const outputFilePath = `output/all_products_${dateTimeString}.docx`;

        //const outputFilePath = path.join('output', 'all_products.docx');
        const buffer = doc.getZip().generate({ type: 'nodebuffer' });

        // Ensure the output directory exists
        fs.mkdirSync(path.dirname(outputFilePath), { recursive: true });

        await writeFileAsync(outputFilePath, buffer);
        console.log(`Generated ${outputFilePath}`);
    } catch (error) {
        console.error("Error creating Word document:", error);
    }
}


// Main function to orchestrate reading CSV, generating content, and creating the document
async function main() {
    try {
        console.log(`Reading file...`);
        const productsCsvPath = 'products.csv'; // Adjust path as needed
        const productsData = await readCsvFile(productsCsvPath);
        console.log(`File reading complete`);

        // Generate responses for each product
        const productsWithContent = await Promise.all(productsData.map(async (product) => {
            const content = await generateResponse(product['name'], product['url']);
            return { ...product, productContent: removeMarkdownFormatting(content) };
        }));
        console.log(`Content generation complete`);

        // Create Word document with the generated content
        await createWordDocument(productsWithContent);
    } catch (error) {
        console.error("Error in main function:", error);
    }
}

// Run the main function
main().catch(console.error);


//Additional functions - Feature upgrades

//remove markdown formatting of content generated by AI
function removeMarkdownFormatting(text) {
    // Remove Markdown headers (### Header)
    let newText = text.replace(/(#+\s*)([^#\n]+)\n/g, '$2\n');
    // Remove Markdown bold syntax (**text**)
    newText = newText.replace(/\*\*(.*?)\*\*/g, '$1');
    return newText;
}

//get current date, time for output filename
function getFormattedDateTime() {
    const now = new Date();
    const day = now.getDate().toString().padStart(2, '0');
    const month = (now.getMonth() + 1).toString().padStart(2, '0'); // Month is 0-indexed
    const year = now.getFullYear();
    const hours = now.getHours().toString().padStart(2, '0');
    const minutes = now.getMinutes().toString().padStart(2, '0');
    const seconds = now.getSeconds().toString().padStart(2, '0');

    return `${day}_${month}_${year}_${hours}_${minutes}_${seconds}`;
}
