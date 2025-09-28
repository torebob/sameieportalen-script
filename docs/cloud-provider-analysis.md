# Cloud Provider Analysis for INFRA-01

## 1. Introduction

This document summarizes the analysis of major cloud providers—Google Cloud Platform (GCP), Microsoft Azure, and Amazon Web Services (AWS)—against the requirements outlined in INFRA-01. The goal is to select a platform that is compliant, supports local data storage, and integrates with Google Workspace.

## 2. Comparison Criteria

*   **GDPR & Norwegian Law:** The platform must document compliance with GDPR and be suitable for storing data under Norwegian jurisdiction.
*   **Data Locality:** The provider must have data centers in Norway or the EU to ensure data residency.
*   **Google Workspace Integration:** The platform must support integration with Google Workspace for core services.

## 3. Platform Comparison

| Feature                       | Google Cloud Platform (GCP)                                  | Microsoft Azure                                             | Amazon Web Services (AWS)                                   |
| ----------------------------- | ------------------------------------------------------------ | ----------------------------------------------------------- | ----------------------------------------------------------- |
| **GDPR/Norwegian Law**        | Strong compliance documentation. Commits to GDPR in contracts. | Strong compliance documentation. Offers tools like Azure Policy to enforce compliance. | Strong compliance documentation. Provides resources and tools for GDPR. |
| **Data Locality (Norway/EU)** | **Yes.** Has a data center region in Oslo, Norway.           | **Yes.** Has data centers in Oslo, Norway.                  | **No (in Norway).** Nearest region is in Sweden. EU regions are available. |
| **Google Workspace Integration** | **Native.** Seamless integration as they are part of the same ecosystem. | **Possible.** Integration is supported via federation (SAML), but requires more complex configuration. | **Possible.** Similar to Azure, integration is achievable through federation but is not native. |

## 4. Summary of Findings

*   **Google Cloud (GCP):** Fully meets all criteria. It offers native integration with Google Workspace, which is a significant advantage, and has a data center in Norway, ensuring data residency and compliance.
*   **Microsoft Azure:** A strong contender. It is fully compliant and has a data center in Norway. However, its integration with Google Workspace is more complex than GCP's.
*   **Amazon Web Services (AWS):** While compliant with GDPR, it lacks a data center in Norway. The nearest region is in Sweden. Integration with Google Workspace is possible but not as straightforward as with GCP.

## 5. Recommendation

Based on the requirements, **Google Cloud Platform (GCP)** is the recommended platform. Its native integration with Google Workspace provides a significant advantage in line with the project's core needs, and its local data center in Oslo satisfies the regulatory requirements for data storage.

While Azure and AWS are viable alternatives, they would introduce additional complexity, particularly regarding Google Workspace integration. GCP offers the most direct and efficient path to meeting all project requirements. This choice also aligns with the user's implicit preference for strong integration with Google services.