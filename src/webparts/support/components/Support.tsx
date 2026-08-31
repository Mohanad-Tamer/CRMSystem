import * as React from 'react';
import styles from './Support.module.scss';

import type { ISupportProps } from './ISupportProps';

interface IFAQ {
  question: string;
  answer: string;
}

const Support: React.FC<ISupportProps> = () => {
  const [openFaq, setOpenFaq] =
    React.useState<number | null>(null);

  const faqs: IFAQ[] = [
    {
      question: 'How do I create a new opportunity?',
      answer:
        'Go to the Sales page from the navigation bar. From there you can view and manage your opportunities and add new opportunity information.'
    },
    {
      question: 'How does an opportunity move to Tendering?',
      answer:
        'An opportunity can move from Sales to Tendering when it is ready for the tendering process. The Tendering page contains the opportunities that have progressed to this stage.'
    },
    {
      question: 'How do I track opportunities in Execution?',
      answer:
        'Opportunities that have been successfully won and handed over for execution can be monitored from the Execution page.'
    },
    {
      question: 'Where can I see the overall CRM performance?',
      answer:
        'Open the Dashboard from the navigation bar. The Dashboard provides pipeline value, opportunity status, domain distribution, pipeline trends and other key performance metrics.'
    },
    {
      question: 'What should I do if the dashboard does not load?',
      answer:
        'First try refreshing the Dashboard. If the problem continues, make sure you have access to the Opportunities SharePoint list and contact your system administrator.'
    }
  ];

  const toggleFaq = (index: number): void => {
    setOpenFaq(
      openFaq === index
        ? null
        : index
    );
  };

  const scrollToFaq = (): void => {
    const faqSection =
      document.getElementById('faq-section');

    if (faqSection) {
      faqSection.scrollIntoView({
        behavior: 'smooth'
      });
    }
  };

  const contactSupport = (): void => {
    window.location.href =
      'mailto:support@example.com';
  };

  const showUserGuide = (): void => {
    window.alert(
      'User Guide will be available here.'
    );
  };

  return (
    <div className={styles.supportPage}>

      {/* =====================================================
          HEADER
          ===================================================== */}

      <div className={styles.supportHeader}>

        <div>

          <div className={styles.eyebrow}>
            CRM SUPPORT
          </div>

          <h1>
            How can we help?
          </h1>

          <p>
            Find answers, guidance, and support for using
            the Opportunities Tracker.
          </p>

        </div>

        <div className={styles.headerIcon}>
          ?
        </div>

      </div>

      {/* =====================================================
          SUPPORT OPTIONS
          ===================================================== */}

      <div className={styles.supportGrid}>

        {/* DOCUMENTATION */}

        <div className={styles.supportCard}>

          <div className={styles.supportCardIconBlue}>
            📖
          </div>

          <div className={styles.supportCardContent}>

            <span className={styles.cardEyebrow}>
              GUIDANCE
            </span>

            <h2>
              User Guide
            </h2>

            <p>
              Learn how to navigate the CRM, manage
              opportunities, and follow the sales process.
            </p>

            <button
              type="button"
              className={styles.supportLink}
              onClick={showUserGuide}
            >
              View User Guide
              <span>
                →
              </span>
            </button>

          </div>

        </div>

        {/* TROUBLESHOOTING */}

        <div className={styles.supportCard}>

          <div className={styles.supportCardIconPurple}>
            🔧
          </div>

          <div className={styles.supportCardContent}>

            <span className={styles.cardEyebrow}>
              TROUBLESHOOTING
            </span>

            <h2>
              Common Issues
            </h2>

            <p>
              Find quick solutions to common problems
              when using the Opportunities Tracker.
            </p>

            <button
              type="button"
              className={styles.supportLinkPurple}
              onClick={scrollToFaq}
            >
              View Solutions
              <span>
                →
              </span>
            </button>

          </div>

        </div>

        {/* CONTACT */}

        <div className={styles.supportCard}>

          <div className={styles.supportCardIconGreen}>
            ✉
          </div>

          <div className={styles.supportCardContent}>

            <span className={styles.cardEyebrow}>
              ASSISTANCE
            </span>

            <h2>
              Contact Support
            </h2>

            <p>
              Need additional help? Contact your CRM
              administrator for assistance.
            </p>

            <button
              type="button"
              className={styles.supportLinkGreen}
              onClick={contactSupport}
            >
              Contact Support
              <span>
                →
              </span>
            </button>

          </div>

        </div>

      </div>

      {/* =====================================================
          QUICK HELP
          ===================================================== */}

      <section className={styles.quickHelp}>

        <div className={styles.sectionHeader}>

          <div>

            <span className={styles.cardEyebrow}>
              QUICK HELP
            </span>

            <h2>
              Getting Started
            </h2>

            <p>
              A quick overview of the main CRM sections.
            </p>

          </div>

        </div>

        <div className={styles.stepsGrid}>

          <div className={styles.stepItem}>

            <div className={styles.stepNumber}>
              01
            </div>

            <div>

              <h3>
                Dashboard
              </h3>

              <p>
                Review your overall pipeline and
                performance metrics.
              </p>

            </div>

          </div>

          <div className={styles.stepItem}>

            <div className={styles.stepNumber}>
              02
            </div>

            <div>

              <h3>
                Sales
              </h3>

              <p>
                Create, review, and manage your
                sales opportunities.
              </p>

            </div>

          </div>

          <div className={styles.stepItem}>

            <div className={styles.stepNumber}>
              03
            </div>

            <div>

              <h3>
                Tendering
              </h3>

              <p>
                Follow opportunities that have
                progressed to tendering.
              </p>

            </div>

          </div>

          <div className={styles.stepItem}>

            <div className={styles.stepNumber}>
              04
            </div>

            <div>

              <h3>
                Execution
              </h3>

              <p>
                Track opportunities that are ready
                for project execution.
              </p>

            </div>

          </div>

        </div>

      </section>

      {/* =====================================================
          FAQ
          ===================================================== */}

      <section
        id="faq-section"
        className={styles.faqSection}
      >

        <div className={styles.sectionHeader}>

          <div>

            <span className={styles.cardEyebrow}>
              FAQ
            </span>

            <h2>
              Frequently Asked Questions
            </h2>

            <p>
              Quick answers to common questions.
            </p>

          </div>

        </div>

        <div className={styles.faqList}>

          {faqs.map(
            (faq, index) => {

              const isOpen =
                openFaq === index;

              return (
                <div
                  key={faq.question}
                  className={
                    isOpen
                      ? styles.faqItemOpen
                      : styles.faqItem
                  }
                >

                  <button
                    type="button"
                    className={styles.faqQuestion}
                    onClick={() =>
                      toggleFaq(index)
                    }
                    aria-expanded={isOpen}
                  >

                    <span>
                      {faq.question}
                    </span>

                    <span
                      className={
                        isOpen
                          ? styles.faqArrowOpen
                          : styles.faqArrow
                      }
                    >
                      {isOpen ? '−' : '+'}
                    </span>

                  </button>

                  {isOpen && (

                    <div className={styles.faqAnswer}>

                      <p>
                        {faq.answer}
                      </p>

                    </div>

                  )}

                </div>
              );
            }
          )}

        </div>

      </section>

      {/* =====================================================
          FOOTER HELP
          ===================================================== */}

      <section className={styles.helpBanner}>

        <div className={styles.helpBannerIcon}>
          ?
        </div>

        <div className={styles.helpBannerText}>

          <h2>
            Still need help?
          </h2>

          <p>
            If you cannot find the answer you are looking
            for, contact your CRM administrator.
          </p>

        </div>

        <button
          type="button"
          className={styles.helpBannerButton}
          onClick={contactSupport}
        >
          Contact Support
        </button>

      </section>

    </div>
  );
};

export default Support;