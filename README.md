# PMP-Charity Matching Analysis System

An intelligent matching system for PMI Sydney Chapter's Project Management Day of Service (PMDoS) that optimally pairs Project Management Professionals (PMPs) with charity organizations based on skills, experience, and professional profiles.

## 🌟 Features

- **Smart Matching Algorithm** - Matches PMPs to charities based on skills, experience, and project requirements
- **LinkedIn Profile Integration** - Validates and scores LinkedIn profiles for professional credibility
- **Comprehensive Reporting** - Multiple output formats including Excel, CSV, and summary reports
- **One-Click Operation** - Single command runs entire analysis pipeline
- **Enhanced Scoring** - Considers professional presence, profile completeness, and skill alignment
- **Ethical LinkedIn Analysis** - URL validation without violating LinkedIn Terms of Service

## 🚀 Quick Start

### Prerequisites

- Python 3.7+
- Required packages: `pandas`, `xlsxwriter`, `openpyxl`

```bash
pip install pandas xlsxwriter openpyxl
```

### Installation

1. Clone this repository:
```bash
git clone https://github.com/joexgit2024/pmi.git
cd pmi
```

2. Place your input Excel files in the `input/` directory:
   - PMP professional registration responses
   - Charity project information responses

3. Run the complete analysis:

   **Standard Assignment (2 PMPs per charity):**
   ```bash
   python run_complete_analysis.py
   ```

   **Flexible Assignment (All PMPs assigned, 3+ PMPs per project):**
   ```bash
   python run_complete_analysis.py --flexible
   ```

### Assignment Modes

- **Standard Mode**: Assigns exactly 2 PMPs per charity project (traditional approach)
- **Flexible Mode**: Assigns all available PMPs to projects with **minimum 2 PMPs per project for risk management**, allowing some projects to have 3+ PMPs for optimal resource utilization

**Risk Management**: Both modes ensure each project has at least 2 PMPs to prevent single points of failure and maintain project continuity.

## 📁 Project Structure

```
pmi/
├── run_complete_analysis.py          # 🎯 MAIN SCRIPT - Only file you need to run
├── enhanced_pmp_charity_matching.py  # Enhanced matching algorithm
├── linkedin_enhanced_matching.py     # LinkedIn profile analysis
├── input/                            # Input Excel files directory
│   └── .gitkeep                      # Keeps directory structure
├── README.md                         # This file
├── .gitignore                        # Git ignore rules
└── requirements.txt                  # Python dependencies
```

## 📊 Input Files

Place these files in the `input/` directory:

1. **PMP Registration Responses** - Excel file containing:
   - PMP professional details
   - Skills ratings (1-5 scale)
   - Experience levels
   - LinkedIn profile URLs
   - Areas of interest

2. **Charity Project Information** - Excel file containing:
   - Organization details
   - Project descriptions
   - Expected outcomes
   - Required skills and expertise

## � Output Files

The system automatically generates:

| File | Description |
|------|-------------|
| `LinkedIn_Analysis_Report.xlsx` | LinkedIn profile validation and quality scoring |
| `PMI_PMP_Charity_Matching_Results_Enhanced.xlsx` | Complete matching results with detailed analysis |
| `Matching_Summary.csv` | Quick reference spreadsheet |
| `Analysis_Summary.txt` | Executive summary |
| `analysis_log.txt` | Detailed processing log |

## � Matching Algorithm

The enhanced matching algorithm considers:

- **Skills Alignment (60%)** - Matches PMP skills with charity project requirements
- **Experience Level (20%)** - Weights experience in project management
- **Interest Alignment (10%)** - Considers PMP's stated areas of interest
- **LinkedIn Quality (5%)** - Professional presence indicator
- **Profile Completeness (5%)** - Completeness of professional information

## 🔧 Configuration

### Customizing the Matching Logic

Edit `enhanced_pmp_charity_matching.py` to modify:
- Skill weighting factors
- Experience scoring
- Project complexity assessment
- Priority algorithms

### LinkedIn Analysis Settings

Modify `linkedin_enhanced_matching.py` to adjust:
- URL validation rules
- Quality scoring criteria
- Profile completeness factors

## 📋 Usage Examples

### Basic Analysis
```bash
python run_complete_analysis.py
```

### Check LinkedIn Profile Quality Only
```bash
python linkedin_enhanced_matching.py
```

### Run Legacy Matching (without LinkedIn)
```bash
python pmp_charity_matching.py
```

## 🛠️ Development

### Key Components

1. **Profile Analysis** - Extracts and scores PMP professional profiles
2. **Requirement Analysis** - Analyzes charity project text for skill requirements
3. **Matching Engine** - Optimally assigns 2 PMPs per charity project
4. **Report Generation** - Creates comprehensive reports and visualizations

### Data Flow

```
Excel Input Files → Profile Analysis → Requirement Analysis → Matching Engine → Reports
                           ↓
                  LinkedIn Analysis → Enhanced Scoring
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🏢 About PMI Sydney Chapter

This system was developed for the PMI Sydney Chapter's Project Management Day of Service (PMDoS), an annual event that connects experienced Project Management Professionals with charitable organizations to provide pro-bono project management expertise.

## 📞 Support

For questions or support, please open an issue in this repository or contact the PMI Sydney Chapter.

---

**Made with ❤️ for the PMI Sydney Chapter community**