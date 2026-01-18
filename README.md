# Project README

## Excel File Content

The Excel file `példatáblázat.xlsx` is a sample evaluation table for a child, containing psychological and educational assessment data.

### Overview
- **File Structure**: 51 rows × 18 columns, with data primarily in column D (labels) and columns E-R (scores/values).
- **Child Name**: "példagyerekneve" (sample child name), located in row 2, column D.
- **Main Sections**: Divided into three key sections marked by headers in column D:
  - **TANSTÍLUS** (row 3): Appears to be a test or category name.
  - **MOTIVÁCIÓ** (row 22): Related to motivation ("motiváció" means motivation in Hungarian).
  - **KATT** (row 39): An abbreviation, possibly a test or category.

### Content Breakdown
- **Learning Style Assessments** (rows 4-8): Evaluations for auditory (auditív), visual (vizuális), and kinesthetic (mozgás) modalities, with subtests and numerical scores (1-5 scale).
- **Additional Categories** (rows 10-14): Assessments for silence (csend), social (társas), meaningful (értelmes), and intuitive (intuitív) traits, with scores.
- **IQ/Test Scores** (row 18): RAVEN test with a score of 45.
- **Other Data**: Sparse numerical values in later rows (e.g., rows 40 and 44 have scores in columns P-R).

The file is mostly empty, serving as a template or partial data entry for child evaluations, with Hungarian labels and a focus on developmental/psychological metrics.

## Technical Articles and Knowledge base

### Various class types...

1. **PO**: (persistent object), persistent object
2. **VO**: (value object), value object
3. **DAO**: (Data Access Objects), data access object interface
4. **BO**: (Business Object), business object layer
5. **DTO** Data Transfer Object data transfer object
6. **POJO**: (Plain Old Java Objects), simple Java objects

The naming conventions with meaningful nomenclature help in understanding the purpose and role of each class in the application architecture. Instead of PO, VO, DAO, BO, DTO, POJO, more descriptive names would enhance code readability and maintainability.

See proposed alternatives following table:

| Abbreviation | Proposed Alternative          | Description                             |
|--------------|-------------------------------|-----------------------------------------|
| PO           | ...Model                      | Persistent Object Model                 |
| VO           | ...       (none)              | Value Object                            |
| DAO          | ...Repository                 | Data access or orm layer                |
| BO           | ...Service or ...Manager      | Business layer service or manager       |
| DTO          | ...Result                     | Returned by the business layer          |
| POJO         | (not specified)               | Plain Old Object                        |
|              | ...Factory                    | Responsible for object creation from various sources |
|              | ...Manager                    | Coordinates multiple objects or operations |

With the use of modern orm libraries and frameworks, some of these patterns may be less relevant, eg. the PO and VO can often be combined into a single model class.
