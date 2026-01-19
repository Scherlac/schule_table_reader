# Models and data structures for Excel processing
import enum
from typing import List, Dict, Tuple, Optional, Any, Annotated
from pydantic import BaseModel, Field


class REPORT_EVAL_ENUM(enum.Flag):
    GRADE_LIMIT = enum.auto()
    HIGH_GRADES = enum.auto()
    COPY = enum.auto()
    MEAN = enum.auto()
    STD = enum.auto()
    MIN = enum.auto()
    MAX = enum.auto()
    SUM = enum.auto()
    NORMED_SUM = enum.auto()

class REPORT_EVAL_TEXT(enum.Enum):
    GRADE_LIMIT = "grade_limit"
    HIGH_GRADES = "> grade_limit"
    COPY = "copy"
    MEAN = "mean"
    STD = "std"
    MIN = "min"
    MAX = "max"
    SUM = "sum"
    NORMED_SUM = "normed sum"

report_eval_to_enum = {
    REPORT_EVAL_TEXT.GRADE_LIMIT.value: REPORT_EVAL_ENUM.GRADE_LIMIT,
    REPORT_EVAL_TEXT.HIGH_GRADES.value: REPORT_EVAL_ENUM.HIGH_GRADES,
    REPORT_EVAL_TEXT.COPY.value: REPORT_EVAL_ENUM.COPY,
    REPORT_EVAL_TEXT.MEAN.value: REPORT_EVAL_ENUM.MEAN,
    REPORT_EVAL_TEXT.STD.value: REPORT_EVAL_ENUM.STD,
    REPORT_EVAL_TEXT.MIN.value: REPORT_EVAL_ENUM.MIN,
    REPORT_EVAL_TEXT.MAX.value: REPORT_EVAL_ENUM.MAX,
    REPORT_EVAL_TEXT.SUM.value: REPORT_EVAL_ENUM.SUM,
    REPORT_EVAL_TEXT.NORMED_SUM.value: REPORT_EVAL_ENUM.NORMED_SUM
}

report_eval_to_text = {v: k for k, v in report_eval_to_enum.items()}


class SectionDetails(BaseModel):
    selection_df: Any = Field(description="The DataFrame containing the selected section data")
    start_row: int = Field(description="The starting row index of the section")
    end_row: int = Field(description="The ending row index of the section")


class RecordDetails(BaseModel):
    source: Any = Field(description="The DataFrame containing the source data")
    location: Tuple[int, int] = Field(description="The (section-relative row, column) coordinates")


class Record(BaseModel):
    label: str = Field(description="The original label string from the Excel row")
    name: str = Field(description="The parsed name of the record")
    items: List[int] = Field(description="List of item numbers associated with the record")
    modifiers: List[str] = Field(description="List of modifiers for each item")
    scores: List[Any] = Field(description="List of score values for the record")
    subsection: Optional[str] = Field(None, description="Optional subsection name if the record belongs to a subsection")
    details: Optional[RecordDetails] = Field(None, description="Details about the source and location of the data")
    eval_results: Optional[Dict[str, Any]] = Field(None, description="Evaluation results for the record")


class SectionConfig(BaseModel):
    approx_start: int = Field(description="Approximate starting row for searching the section")
    multi_row: bool = Field(description="Whether records in this section span multiple rows")
    expected_records: Optional[int] = Field(None, description="Expected number of records in the section")
    expected_questions: Optional[List[int]] = Field(None, description="Expected number of questions per record")
    # "classification": [1, 1, 1, 1, 1, 2, 2, 3, 3],
    classification: Optional[List[int]] = Field(None, description="Class identifiers for each record")
    # "class_marker": ["cognitive", "social", "logic"],
    class_marker: Optional[List[str]] = Field(None, description="Class markers for grouping records")
    # "question_id": [11, 12, 21, 22, 31, 41, 42, 51, 52],
    question_id: Optional[List[int]] = Field(None, description="Question IDs for each record")
    # "report_eval": "> max - dev"
    class_eval: Optional[str] = Field(None, description="Evaluation criteria for reporting scores on class level")
    record_eval: Optional[str] = Field(None, description="Evaluation criteria for reporting scores on record level")

    @property
    def class_eval_flags(self) -> REPORT_EVAL_ENUM:
        """
        Parses the class_eval string into a combined REPORT_EVAL_ENUM flag.

        Returns:
        REPORT_EVAL_ENUM: Combined flags based on the class_eval configuration.
        """
        if not self.class_eval:
            return REPORT_EVAL_ENUM(0)
        flags = REPORT_EVAL_ENUM(0)
        parts = [part.strip() for part in self.class_eval.split(',')]
        for part in parts:
            if part in report_eval_to_enum:
                flags |= report_eval_to_enum[part]
        return flags

    @class_eval_flags.setter
    def class_eval_flags(self, flags: REPORT_EVAL_ENUM) -> None:
        """
        Sets the class_eval string based on the provided REPORT_EVAL_ENUM flags.

        Parameters:
        flags (REPORT_EVAL_ENUM): Combined flags to set the class_eval configuration.
        """
        parts = []
        for key, value in report_eval_to_enum.items():
            if flags & value:
                parts.append(key)
        self.class_eval = ', '.join(parts)

    @property
    def record_eval_flags(self) -> REPORT_EVAL_ENUM:
        """
        Parses the record_eval string into a combined REPORT_EVAL_ENUM flag.

        Returns:
        REPORT_EVAL_ENUM: Combined flags based on the record_eval configuration.
        """
        if not self.record_eval:
            return REPORT_EVAL_ENUM(0)
        flags = REPORT_EVAL_ENUM(0)
        parts = [part.strip() for part in self.record_eval.split(',')]
        for part in parts:
            if part in report_eval_to_enum:
                flags |= report_eval_to_enum[part]
        return flags

    @record_eval_flags.setter
    def record_eval_flags(self, flags: REPORT_EVAL_ENUM) -> None:
        """
        Sets the record_eval string based on the provided REPORT_EVAL_ENUM flags.

        Parameters:
        flags (REPORT_EVAL_ENUM): Combined flags to set the record_eval configuration.
        """
        parts = []
        for key, value in report_eval_to_enum.items():
            if flags & value:
                parts.append(key)
        self.record_eval = ', '.join(parts)

class ClassResult(BaseModel):
    class_id: int = Field(description="Identifier for the class")
    marker: str = Field(description="Marker for the class")
    records: List[Record] = Field(description="List of records belonging to the class")
    eval_results: Dict[str, Any] = Field(description="Evaluation results for the class")

class SectionResult(BaseModel):
    details: SectionDetails = Field(description="Details about the section extraction")
    records: List[Record] = Field(description="List of parsed records from the section")
    class_results: Optional[Dict[str, ClassResult]] = Field({}, description="Optional mapping of class markers to their results")
    # subsections: Optional[Dict[str, List[Record]]] = Field(default=None, description="Optional mapping of subsection names to their records")