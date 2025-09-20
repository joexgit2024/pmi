# Setup Guide for PMI-Charity Matching System

## 🚀 Quick Setup

### 1. Clone the Repository
```bash
git clone https://github.com/joexgit2024/pmi.git
cd pmi
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Prepare Input Data
1. Create your Excel input files with PMP and charity data
2. Place them in the `input/` directory
3. Ensure file names match the expected format (see README.md)

### 4. Run Analysis
```bash
python run_complete_analysis.py
```

## 📋 Expected Input File Format

### PMP Registration File
Should contain columns:
- First Name, Last Name
- Email address
- LinkedIn Profile URL
- Current / Latest Job Title
- Company
- Year(s) as a Project Professional
- Areas of Interest
- Skill ratings (1-5 scale) for various PM competencies

### Charity Project File
Should contain columns:
- Name of the organisation
- Name of the initiative
- Simple description of the initiative or project
- Key outcomes expected
- How will this benefit the organisation
- Expected outcomes from PMDoS participation

## 🔧 Customization

- Modify matching weights in `enhanced_pmp_charity_matching.py`
- Adjust LinkedIn scoring in `linkedin_enhanced_matching.py`
- Update skill categories as needed
- Customize report formats

## 📞 Support

Open an issue on GitHub if you encounter any problems or need assistance with setup.