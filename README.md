# Automated Document Distribution  
### Excel + Google Apps Script Integration  
#### vibecoding with ChatGPT  

---

## What this is

A small but fully working automation that takes boring repetitive work —  
copying names, generating PDFs, sending dozens of emails —  
and makes Excel and Gmail do it *for you* (mostly).

Built by a human with **zero formal programming background**,  
and a lot of patience, curiosity, and ChatGPT on speed dial.  

---

## What it does

This workflow:
- Reads data from an Excel (or Google Sheet) file  
- Generates personalized PDFs based on templates (V1, V2, V3)  
- Sends individual emails through Gmail using Google Apps Script  
- Keeps everything human-controlled and privacy-safe (sandboxed mode)  

In short:  
> *It automates what you already do manually, without losing control of the send button.*

---

## Tech stack (a.k.a. “tools I didn’t know existed two weeks ago”)

- **Microsoft Excel VBA** – for data parsing and PDF generation  
- **Google Apps Script (Gmail API)** – for sending customized emails  

---

## Why I built it

I work in a commercial back office.  
At some point, sending 20+ personalized communications manually felt like a crime against productivity.  

So I decided to see if I could automate the whole process **without being a developer** —  
just me, some templates, and a chatbot with endless patience.  

This project is the result of learning one concept at a time —  
or, as I like to call it:  
> **Learning one mess at a time™.**

---

## How it works (simplified)

1. **Excel side:**  
   - Reads customer data  
   - Chooses the right Word template  
   - Generates a PDF per row  

2. **Google side:**  
   - Reads the same data  
   - Sends the corresponding email via Gmail  
   - Logs everything (because chaos is not a strategy)

3. **Human side:**  
   - Braces herself for the worse  
   - Clicks *Run*  
   - Smiles when it actually works

---

## What I learned

- How to break a problem into steps instead of breaking my keyboard  
- That AI isn’t magic — it’s a very patient co-pilot (if you know what to ask)  
- Debugging is 80% detective work and 20% internal screaming  
- And that automation feels *really* good when it finally runs

---

## Acknowledgments

- **ChatGPT (OpenAI)** – my 24/7 coding buddy and therapist  
- **Google Apps Script** – for existing in a weirdly friendly way  
- **Microsoft Excel** – for being both the problem and the solution  
- **Me** – for not giving up when the first 12 versions didn’t work  

---

## License

Do whatever you want with it.  
Just don’t send spam, don’t steal my vibe, and please credit *Learning one mess at a time™*.

---

## Final thought

This repo isn’t about being a developer.  
It’s about learning by *doing the thing anyway*.  

If I can automate document workflows with a chatbot and some curiosity,  
you probably can too.  

