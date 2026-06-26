# Project context

This is my agent workspace. I will use this space to upload many different types of data formats and I would like these formats to be standardised into a single clean format. Some of the data formats will have errors, different column names and missing data. We need to create a well defined set of rules to correct the errors, standardise the column names and fill in the missing data. Note there will be 100+ different input raw formats

# About me

I prefer clean, jargon free output

# Rules

⁃ Always ask clarifying questions before starting a complex task 
- Show your plan and steps before executing anything
- Keep reports and summaries concise 
- bullet points over paragraphs 
- Save all output files to the output folder 
- do not make changes until you have 95% confidence on what you going to build. Asked me follow up question until you reach 95% confidence.
- Only care about KWh or MWh (not KVA or KVAH)
- Consume Energy (CE) = KWh consumption or sometimes denoted as E
- Sent Out Energy (SOE) = KWh of export or sometimes denoted as B i.e solar export
- Net Energy =  CE - absolute (SOE)
- NMI = 10 characters or 11 or 12 in the input files
- Sent Out Energy (SOE) should always be a negative number. If its positive, multiply by -1
- If NMI is more than 10 characters in the input file, then just take the first 10 characters
- clean output file must be .csv
- Inputs files can be in 5 minute, 15 minute, 30 minute and/or 60 minute time intervals. In any case, please convert into 30 minute as the standard.

# Project Structure

- workflows/ - workflow instruction files (plain english recipes the agent follows)
- output/ - finish deliverables (reports, drafts, analysis)
- Resources/ - reference docs and templates