# DMart Sales & Profit Analysis

## Executive Summary

### Overall Performance

| Metric | Value |
|---|---:|
| Total Records | 15,346 |
| Total Sales | ₹3,92,62,877.52 |
| Total Profit | ₹68,72,139.81 |
| Average Order Value (AOV) | ₹1,915.19 |
| Total Quantity Sold | 76,831 |
| Average Quantity Sold | 5 |

### Key Insights

- Profit Margin ≈ **17.5%** (68.72L / 392.63L)
- Average order value is healthy at **₹1,915**.
- Sales and profits are distributed fairly evenly across regions and cities.
- No region is significantly underperforming.
- Chennai delivers the highest city profit.
- North region generates the highest overall sales and profit.

---

## Region-wise Analysis

| Region | Total Profit | Total Sales | AOV |
|---|---:|---:|---:|
| East | ₹16.75L | ₹95.18L | ₹443.43 |
| North | ₹17.69L | ₹101.44L | ₹450.82 |
| South | ₹17.14L | ₹98.52L | ₹444.86 |
| West | ₹17.14L | ₹97.49L | ₹452.19 |

### Findings

#### Best Performing Region — North

- Highest Profit = ₹17.69L
- Highest Sales = ₹101.44L

#### Highest AOV — West

- ₹452.19

This indicates customers in West spend slightly more per order.

#### Lowest Profit Transaction — North

- ₹5.76

#### Highest Single Transaction Profit — North

- ₹1,486.66

### Recommendation

Focus marketing efforts on North because:

- Highest sales
- Highest profit
- Strong transaction profitability

---

## Category-wise Analysis

| Category | Total Profit | Total Sales | AOV |
|---|---:|---:|---:|
| Clothing | ₹13.38L | ₹76.77L | ₹445.66 |
| Electronics | ₹13.99L | ₹80.18L | ₹449.09 |
| Home Care | ₹13.94L | ₹79.29L | ₹447.63 |
| Grocery | ₹13.51L | ₹77.19L | ₹446.57 |
| Personal Care | ₹13.90L | ₹79.20L | ₹450.17 |

### Findings

#### Highest Profit — Electronics

- ₹13.99L

#### Highest Sales — Electronics

- ₹80.18L

#### Highest AOV — Personal Care

- ₹450.17

#### Highest Single Transaction Profit — Electronics

- ₹1,486.66

#### Lowest Profit Transaction — Personal Care

- ₹5.76

### Recommendation

Electronics is the strongest category:

- Highest revenue
- Highest profit
- Strong margins

Personal Care customers spend more per transaction, suggesting opportunities for premium products.

---

## City-wise Analysis

| City | Total Profit | Total Sales | AOV |
|---|---:|---:|---:|
| Bangalore | ₹11.50L | ₹66.05L | ₹441.19 |
| Chennai | ₹11.65L | ₹66.44L | ₹454.13 |
| Delhi | ₹11.59L | ₹66.13L | ₹452.83 |
| Mumbai | ₹11.58L | ₹65.64L | ₹455.43 |
| Pune | ₹11.19L | ₹63.09L | ₹440.83 |
| Hyderabad | ₹11.21L | ₹64.47L | ₹442.69 |

### Findings

#### Highest Profit — Chennai

- ₹11.65L

#### Highest Sales — Chennai

- ₹66.44L

#### Highest AOV — Mumbai

- ₹455.43

#### Lowest Sales — Pune

- ₹63.09L

#### Lowest Profit — Hyderabad

- ₹11.21L

### Recommendation

- Chennai is the most profitable city.
- Mumbai customers spend the most per order.
- Pune may require promotional campaigns to increase revenue.

---

## Profitability Analysis

### Profit Margin

Formula:

```excel
= Total Profit / Total Sales
```

Example:

```excel
=6872139.81/39262877.52
```

Result:

```text
17.50%
```

This is a strong retail margin.

---

# Excel Formulas Used

Assume:

| Column | Data |
|---|---|
| A | Order ID |
| B | Region |
| C | Category |
| D | City |
| E | Sales |
| F | Profit |
| G | Quantity |

## 1. Total Records

```excel
=COUNTA(A:A)-1
```

or

```excel
=ROWS(A2:A15347)
```

## 2. Total Sales

```excel
=SUM(E:E)
```

## 3. Total Profit

```excel
=SUM(F:F)
```

## 4. Total Quantity Sold

```excel
=SUM(G:G)
```

## 5. Average Order Value (AOV)

```excel
=AVERAGE(E:E)
```

or

```excel
=SUM(E:E)/COUNT(E:E)
```

## 6. Average Quantity Sold

```excel
=AVERAGE(G:G)
```

---

# Region-wise Formulas

Assume region name is in K2.

### Count

```excel
=COUNTIF($B:$B,K2)
```

### Total Profit

```excel
=SUMIFS($F:$F,$B:$B,K2)
```

### AOV

```excel
=AVERAGEIFS($E:$E,$B:$B,K2)
```

### Max Profit

```excel
=MAXIFS($F:$F,$B:$B,K2)
```

### Min Profit

```excel
=MINIFS($F:$F,$B:$B,K2)
```

### Total Sales

```excel
=SUMIFS($E:$E,$B:$B,K2)
```

### Average Sales

```excel
=AVERAGEIFS($E:$E,$B:$B,K2)
```

---

# Category-wise Formulas

Assume category name in K20.

### Count

```excel
=COUNTIF($C:$C,K20)
```

### Total Profit

```excel
=SUMIFS($F:$F,$C:$C,K20)
```

### AOV

```excel
=AVERAGEIFS($E:$E,$C:$C,K20)
```

### Max Profit

```excel
=MAXIFS($F:$F,$C:$C,K20)
```

### Min Profit

```excel
=MINIFS($F:$F,$C:$C,K20)
```

### Total Sales

```excel
=SUMIFS($E:$E,$C:$C,K20)
```

### Average Sales

```excel
=AVERAGEIFS($E:$E,$C:$C,K20)
```

---

# City-wise Formulas

Assume city name in K40.

### Count

```excel
=COUNTIF($D:$D,K40)
```

### Total Profit

```excel
=SUMIFS($F:$F,$D:$D,K40)
```

### AOV

```excel
=AVERAGEIFS($E:$E,$D:$D,K40)
```

### Max Profit

```excel
=MAXIFS($F:$F,$D:$D,K40)
```

### Min Profit

```excel
=MINIFS($F:$F,$D:$D,K40)
```

### Total Sales

```excel
=SUMIFS($E:$E,$D:$D,K40)
```

### Average Sales

```excel
=AVERAGEIFS($E:$E,$D:$D,K40)
```

---

# Dashboard Highlights (for presentation)

1. **North Region** contributes the highest sales (₹1.01 Cr) and profit (₹17.69 L).
2. **West Region** has the highest average order value (₹452.19).
3. **Electronics** is the best-performing category by sales and profit.
4. **Personal Care** has the highest category AOV.
5. **Chennai** leads among cities in both sales and profit.
6. **Mumbai** customers spend the most per order.
7. Overall profit margin is approximately **17.5%**, indicating healthy business performance.
8. Sales distribution is balanced across regions, reducing dependency on a single market.
9. Pune and Hyderabad present opportunities for targeted growth initiatives.
10. Electronics and Personal Care should be prioritized for expansion and promotional campaigns.
