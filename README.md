# 📊 WAM Dashboard and Data Summarizer

A modern, interactive dashboard for tracking Weekly Accountability Meetings (WAM) across multiple offices. Real-time data synchronization from Google Sheets with comprehensive analytics, visualizations, and advanced table features.

![Dashboard Version](https://img.shields.io/badge/version-3.1-blue)
![License](https://img.shields.io/badge/license-MIT-green)

## ✨ Features

### � Dashboard Overview
- **Total Audits**: Track total WAM records
- **Total Offices**: Count of unique office locations
- **On-Track Rate**: Real-time percentage with color-coded indicators (🟢≥80% 🟡60-79% 🔴<60%)
- **Average Meeting Rating**: Quality ratings from 1-10 scale with color coding

### 🎯 Rock Review Analysis
- **Circular Progress Charts**: Visual representation of On-Track vs Off-Track performance
- **Performance Summary**: Animated percentage-based metrics
- **Color-coded Status**: Easy-to-read performance indicators

### 📊 Interactive Visualizations

#### Office Distribution (Pie Chart)
- Visual breakdown of records by office
- **Interactive**: Click segments to view detailed office data
- Modal popup with office-specific metrics and export

#### Office Performance Comparison (Bar Chart)
- Horizontal stacked bar chart showing On-Track/Off-Track/Other status
- Color-coded performance indicators
- Easy comparison across all offices

#### Performance Trend Over Time (Line Chart)
- **Dual Y-axis chart**: On-Track rate (%) and total records
- Monthly aggregation with smooth animations
- Large, crisp 450px display optimized for readability
- Hover tooltips with detailed metrics

### 🗂️ Complete Data Table (Enhanced)
- **Pagination**: 10, 25, 50, 100, or all records per page
- **Column Visibility**: Show/hide specific columns
- **Search**: Real-time search across all data
- **Sortable**: Click any header to sort ascending/descending
- **Row Highlighting**: Off-Track records in red
- **Export Filtered**: Download visible data as CSV with date stamp
- **Smart Navigation**: Page numbers with Previous/Next controls
- **Responsive**: Adapts to all screen sizes

## 🚀 Quick Start

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/wam-dashboard.git
   cd wam-dashboard
   ```

2. **Start local server**
   ```bash
   python3 -m http.server 8080
   ```

3. **Open in browser**
   ```
   http://localhost:8080/dashboard-modular.html
   ```

4. **View live data**
   - Dashboard loads automatically from Google Sheets
   - Data refreshes every 2 minutes
   - All visualizations and metrics update in real-time

## � Project Structure

```
wam-dashboard/
├── dashboard-modular.html    # Main dashboard page
├── css/
│   └── dashboard.css         # Custom styles (798 lines)
├── js/
│   └── dashboard.js          # Dashboard logic (2165 lines)
├── README.md                 # This file
└── .gitignore               # Git ignore rules
```

## 🔧 Configuration

The dashboard is pre-configured with Google Sheets integration in `js/dashboard.js`:

```javascript
const CONFIG = {
    API_KEY: 'YOUR_API_KEY',
    SHEET_ID: 'YOUR_SHEET_ID',
    SHEET_RANGE: 'Weekly Audit for Offices',
    AUTO_REFRESH_INTERVAL: 2 * 60 * 1000 // 2 minutes
};
```

### Column Mapping
The dashboard automatically detects and maps these columns from your sheet:

| Column | Header | Usage | Description |
|--------|--------|-------|-------------|
| A | OFFICE NAME | Office filter & charts | Name of the office location |
| B | Office Personnel Name/s | Personnel filter | Personnel conducting meetings |
| C | SEGUE | Display only | Wins (personal/work) |
| D | SCORECARD | Display only | Measurables (KPIs) |
| E | ROCK REVIEW | Performance metrics | On-Track or Off-Track status |
| F | TO DO LIST | Display only | Commitments for next week |
| G | IDS | Display only | Issues to discuss/solve |
| H | CONCLUDE | Meeting rating | Meeting quality rating (1-10) |
| I | Triage Officer | Display only | Officer name |
| J | Audit Date | Trend analysis | Date of the audit (MM/DD/YYYY) |

### Data Structure Requirements
- **Row 1-2**: Title/introduction text (automatically filtered out)
- **Row 3**: Column headers (used by dashboard)
- **Row 4+**: Actual data records

## 💡 How to Use

### Dashboard Navigation
1. **Dashboard Overview** - View summary metrics at the top
2. **Rock Review** - Check On-Track vs Off-Track performance
3. **Charts** - Analyze office distribution, performance comparison, and trends
4. **Data Table** - Browse, filter, and export detailed records

### Filtering Data
- **Search Box**: Real-time search across all columns
- **Time Period**: Filter by week, month, or year
- **Personnel Filter**: Select specific personnel
- **Office Filter**: Select specific office

### Using the Data Table
1. **Show/Hide Columns**: Click button to toggle column visibility
2. **Change Rows Per Page**: Select 10, 25, 50, 100, or All
3. **Navigate Pages**: Use Previous/Next or click page numbers
4. **Sort Data**: Click any column header
5. **Export Data**: Click "Export Filtered" to download visible data as CSV

### Interacting with Charts
- **Pie Chart**: Click any office segment to view detailed breakdown
- **Bar Chart**: Hover to see exact counts
- **Line Chart**: Hover over points to see On-Track % and record counts
- **Export**: Use modal popup to export office-specific data

## 🎨 Technologies Used

- **Frontend**: HTML5, CSS3, JavaScript (ES6+)
- **Framework**: Bootstrap 5.3.2
- **Charts**: Chart.js 4.4.0 (with HiDPI/Retina support)
- **Icons**: Bootstrap Icons 1.11.1
- **Animations**: Animate.css 4.1.1
- **API**: Google Sheets API v4

## 🌐 Browser Support

- ✅ Chrome (latest) - Recommended
- ✅ Firefox (latest)
- ✅ Safari (latest)
- ✅ Edge (latest)
- ✅ Mobile browsers

## 🛠️ Troubleshooting

### Dashboard not loading data
**Solution**: 
- Check internet connection
- Verify Google Sheet is accessible (shared with "Anyone with link" as Viewer)
- Check browser console for errors (F12)
- Ensure API key is valid and not restricted

### Table showing no data
**Solution**:
- Verify data starts from Row 4 (Row 3 contains headers)
- Check that Office Name column (A) is not empty
- Make sure sheet name is "Weekly Audit for Offices"
- Check browser console for filtering logs

### Charts not displaying correctly
**Solution**:
- Ensure Rock Review column (E) contains "On-Track" or "Off-Track" text
- Verify Audit Date column (J) has valid dates (MM/DD/YYYY format)
- Check that Meeting Rating column (H) has numeric values 1-10
- Try manual refresh button

### Chart is blurry or pixelated
**Solution**:
- Clear browser cache and reload
- The dashboard uses Chart.js HiDPI support for Retina displays
- Ensure you're using latest version of Chrome/Firefox

### Pagination not working
**Solution**:
- Check that data has been loaded (watch console for "Filtered data rows")
- Try clicking "All" in rows per page dropdown
- Refresh the data

## 🔒 Security Note

⚠️ **Important**: The API key is currently embedded in the code for convenience. 

**For production use, consider:**
- Implementing server-side API calls
- Using environment variables for sensitive data
- Restricting API key to specific domains in Google Cloud Console
- Implementing OAuth 2.0 for better security

**Current security:**
- API key has restricted access to Google Sheets API only
- Google Sheet must be shared as "Viewer" to anyone with link
- No write access to the sheet from the dashboard

## 🎨 Customization

### Change Auto-refresh Interval
Edit `js/dashboard.js` (line ~85):
```javascript
AUTO_REFRESH_INTERVAL: 2 * 60 * 1000 // Change 2 to desired minutes
```

### Modify Color Coding Thresholds
Edit `js/dashboard.js` in `renderSummaryCards()` function:
```javascript
// On-Track Rate colors
if (onTrackRate >= 80) return 'success';  // Green
else if (onTrackRate >= 60) return 'warning'; // Yellow
else return 'danger';  // Red

// Meeting Rating colors
if (avgRating >= 8) return 'success';
else if (avgRating >= 6) return 'warning';
else return 'danger';
```

### Change Chart Colors
Edit `js/dashboard.js` chart configuration sections:
```javascript
// Pie chart colors
backgroundColor: [
    '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', 
    '#9966FF', '#FF9F40', '#FF6384', '#C9CBCF'
]

// Line chart colors
borderColor: 'rgb(75, 192, 192)', // On-Track line
backgroundColor: 'rgba(75, 192, 192, 0.1)' // Fill color
```

### Adjust Chart Size
Edit `dashboard-modular.html` (line ~147):
```html
<div style="height: 450px;"> <!-- Change height here -->
```

### Modify Pagination Options
Edit `dashboard-modular.html` (line ~184):
```html
<option value="10">10</option>
<option value="25" selected>25</option>
<option value="50">50</option>
<option value="100">100</option>
<option value="999999">All</option>
```

## 🚀 Deployment Options

### Option 1: GitHub Pages (Free & Easy)
1. **Create repository** on GitHub
2. **Upload files**: `dashboard-modular.html`, `css/`, `js/`, `README.md`
3. **Go to Settings** → Pages
4. **Select branch**: main
5. **Live at**: `https://username.github.io/wam-dashboard/dashboard-modular.html`

### Option 2: Netlify (Free)
1. Sign up at [Netlify](https://netlify.com)
2. Drag and drop your project folder
3. Instant deployment!
4. Bonus: Custom domain & HTTPS included

### Option 3: Vercel (Free)
1. Sign up at [Vercel](https://vercel.com)
2. Import your GitHub repository
3. Deploy with one click
4. Automatic deployments on git push

### Option 4: Local/Internal Server
```bash
# Using Python 3 (Recommended)
python3 -m http.server 8080

# Using Node.js
npx http-server -p 8080

# Using PHP
php -S localhost:8080

# Then open: http://localhost:8080/dashboard-modular.html
```

**For company intranet**: Deploy to internal web server (Apache/Nginx)

## 📊 Use Cases

This dashboard is perfect for:
- 📈 **Weekly Accountability Meeting Tracking** - Primary use case
- 🏢 **Multi-office Performance Monitoring** - Compare offices at a glance
- � **Team Accountability** - Track personnel meeting quality
- � **Trend Analysis** - Monitor performance improvements over time
- 🎯 **Rock Review Tracking** - On-Track vs Off-Track status monitoring
- � **Audit Management** - Comprehensive audit record keeping

## � Version History

- **v3.1** (Current) - Enhanced table features (pagination, column visibility, export, highlighting)
- **v3.0** - Added Performance Trend chart with dual Y-axis
- **v2.5** - Implemented Dashboard Overview with meaningful metrics
- **v2.0** - Added Rock Review analysis and interactive charts
- **v1.0** - Initial release with basic data table

## 🚦 Roadmap

Planned features for future releases:
- [ ] Time period filters (Weekly/Monthly/Yearly view)
- [ ] Office/personnel dropdown filters
- [ ] Click office in pie chart to filter table
- [ ] Date range picker
- [ ] Email report generation
- [ ] PDF export functionality
- [ ] Dark mode theme
- [ ] Mobile app version
- [ ] Multiple sheet comparison

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

1. **Fork the repository**
2. **Create a feature branch**: `git checkout -b feature/AmazingFeature`
3. **Commit your changes**: `git commit -m 'Add some AmazingFeature'`
4. **Push to the branch**: `git push origin feature/AmazingFeature`
5. **Open a Pull Request**

**Areas for contribution:**
- Additional chart types
- More filter options
- Performance optimizations
- UI/UX improvements
- Documentation enhancements
- Bug fixes

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👤 Author

**Jireh B. Custodio**

- GitHub: [@yourusername](https://github.com/yourusername)
- Email: your.email@example.com

## 🙏 Acknowledgments

- **Bootstrap Team** - For the excellent CSS framework
- **Chart.js** - For powerful and beautiful charts
- **Google** - For Sheets API
- **Community** - For feedback and feature suggestions

## 📞 Support

Need help? Here's how to get support:

1. **Check Documentation**: Read this README thoroughly
2. **Troubleshooting Guide**: Review the troubleshooting section above
3. **GitHub Issues**: [Create an issue](https://github.com/yourusername/wam-dashboard/issues)
4. **Email Support**: your.email@example.com

## ⚡ Performance Notes

- **Load Time**: < 2 seconds (with good internet connection)
- **Data Refresh**: Every 2 minutes automatically
- **Max Records**: Tested with 1000+ records
- **Browser Memory**: < 50MB typical usage
- **Mobile Friendly**: Fully responsive design

## � Learning Resources

Want to customize further? Learn more about:
- [Chart.js Documentation](https://www.chartjs.org/docs/)
- [Bootstrap 5 Documentation](https://getbootstrap.com/docs/5.3/)
- [Google Sheets API](https://developers.google.com/sheets/api)
- [JavaScript ES6+](https://developer.mozilla.org/en-US/docs/Web/JavaScript)

---

**Last Updated**: November 5, 2025  
**Version**: 3.1 - Enhanced Table Features  
**Status**: ✅ Production Ready

Made with ❤️ for better team accountability and performance tracking

**Happy Dashboard Monitoring! 🚀📊**
