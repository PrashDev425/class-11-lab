# Lab: Mail Merge in MS Word

## Task

You are working in a school office. You need to send a personalized invitation letter to parents of students for the **Annual Parent Meeting**. Use **Mail Merge** with a CSV file to create and send the letters.

---

## Part A – Create the Data Source (CSV)

**Data Sample for parents:**

|FirstName|LastName|ParentName        |Email           |Address      |
|---------|--------|------------------|----------------|-------------|
|Ravi     |Sharma  |Mr. Sunil Sharma  |sunil@gmail.com |New Baneshwor|
|Anisha   |Koirala |Mrs. Meena Koirala|meena@yahoo.com |Lalitpur     |
|Sita     |Thapa   |Mr. Hari Thapa    |hari@hotmail.com|Bhaktapur    |

1. Open **Notepad** or any text editor.
2. Enter the following data:

    ```csv
    FirstName,LastName,ParentName,Email,Address
    Ravi,Sharma,Mr. Sunil Sharma,sunil@gmail.com,New Baneshwor
    Anisha,Koirala,Mrs. Meena Koirala,meena@yahoo.com,Lalitpur
    Sita,Thapa,Mr. Hari Thapa,hari@hotmail.com,Bhaktapur
    ```

3. Save the file as `Parents.csv`.
4. Close the file.

---

## Part B – Create the Main Document (Word)

1. Open **MS Word** → New Blank Document.
2. Go to **Mailings → Start Mail Merge → Letters** (or **E-Mail Messages**).
3. Connect to your CSV file:

   * **Select Recipients → Use an Existing List → Parents.csv**.
   * Make sure the first row contains column headers.

---

## Part C – Write the Letter

**Letter Sample**

```
Dear «ParentName»,

This is to invite you to the Annual Parent Meeting for your ward «FirstName» «LastName».
The meeting will be held at the School Hall, Kathmandu on 15th September at 10:00 AM.

We kindly request your presence.

Sincerely,  
Principal
```

Type the following letter in Word and insert merge fields:

* Insert merge fields via **Mailings → Insert Merge Field**.

---

## Part D – Preview & Finish

1. Click **Preview Results** to check the letters.
2. Go to **Finish & Merge → Edit Individual Documents** to create the final letters.
3. Save the merged letters as `ParentInvitations.docx`.

---

## Questions & Answers

**Q1. What is Mail Merge used for?**
**A:** Mail Merge is used to create many personalized documents (letters, emails, labels, envelopes) at once by combining a main document with a data source.

**Q2. What are the two main parts of Mail Merge?**
**A:**

1. **Main Document** (the letter or email template)
2. **Data Source** (list of names, addresses, etc., usually in CSV or Excel)

**Q3. How do you insert a field like First Name in Word?**
**A:** Place the cursor where you want it → go to **Mailings → Insert Merge Field → FirstName**.

**Q4. Can Mail Merge be used for email? If yes, how?**
**A:** Yes. Choose **E-Mail Messages** in Mail Merge. Then use **Finish & Merge → Send E-Mail Messages**, select the **Email** field, type a subject, and send via Outlook.

---

## Deliverables

* `Parents.csv` (data source)
* `ParentInvitations.docx` (merged letters)
* Written answers to the questions
