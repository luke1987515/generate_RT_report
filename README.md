# Generate RT Report

[English](README.md) | [繁體中文](README.zh-TW.md)

This repository contains a collection of scripts and tools for automated Microsoft Word (RT) report generation and manipulation. 

## Features

The project is divided into several sub-modules targeting different aspects of report automation:

* **Automated Word Report Generation**: Scripts for generating reports with specific formatting and structure (`generate_RT_report_*`).
* **Table Modification & Timestamps**: Tools for modifying tables within Word documents and logging/updating dates (`auto_report-table-date_piero`, `auto_report-table_modify_piero`).
* **Outline Processing**: Utilities for reading and filtering Word document outlines (`read_word_outline`, `word_outline_filter`).
* **Template Management**: Converting documents to templates (`doc_to_template`).
* **Integrated Tools**: An integrated report tool combining various features (`integrated_report_tool`).

## Structure

* `auto_report-table-date_piero/`: Scripts for date and time management in report tables.
* `auto_report-table_modify_piero/`: Scripts for general table modifications.
* `doc_to_template/`: Document to template conversion utilities.
* `find_tables_in_word/`: Tooling to locate and extract tables from Word documents.
* `generate_RT_report_luke/` & `generate_RT_report_piero/`: Core report generation scripts by different contributors.
* `integrated_report_tool/`: The integrated tool combining multiple features.
* `read_word_outline/` & `word_outline_filter/`: Scripts for handling Word outlines and sections.

## Getting Started

Each directory contains specific tools and scripts. Please navigate to the respective directory for more detailed information or scripts related to that specific task.
