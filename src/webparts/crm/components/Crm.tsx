import * as React from 'react';
import styles from './Crm.module.scss';

import Opportunities from '../../opportunities/components/Opportunities';
import Tendering from '../../tendering/components/Tendering';
import Execution from '../../execution/components/Execution';
import Dashboard from '../../dashboard/components/Dashboard';
import Support from '../../support/components/Support';

import type { ICrmProps } from './ICrmProps';

type Page =
  | 'home'
  | 'dashboard'
  | 'sales'
  | 'tendering'
  | 'execution'
  | 'support';

const Crm: React.FC<ICrmProps> = (props) => {
  const [currentPage, setCurrentPage] =
    React.useState<Page>('home');

  const renderPage = (): React.ReactElement => {
    switch (currentPage) {

      /* =====================================================
         DASHBOARD
         ===================================================== */

      case 'dashboard':
        return (
          <div className={styles.page}>
            <Dashboard context={props.context} />
          </div>
        );

      /* =====================================================
         SALES
         ===================================================== */

      case 'sales':
        return (
          <div className={styles.page}>
            <div className={styles.pageTitle}>
              <h1>Sales</h1>
              <p>
                Manage your sales opportunities.
              </p>
            </div>

            <Opportunities context={props.context} />
          </div>
        );

      /* =====================================================
         TENDERING
         ===================================================== */

      case 'tendering':
        return (
          <div className={styles.page}>
            <div className={styles.pageTitle}>
              <h1>Tendering</h1>

              <p>
                Manage opportunities that have moved
                from Sales into Tendering.
              </p>
            </div>

            <Tendering context={props.context} />
          </div>
        );

      /* =====================================================
         EXECUTION
         ===================================================== */

      case 'execution':
        return (
          <div className={styles.page}>
            <div className={styles.pageTitle}>
              <h1>Execution Tracker</h1>

              <p>
                Track opportunities that were won
                during Tendering and are ready for execution.
              </p>
            </div>

            <Execution context={props.context} />
          </div>
        );

      /* =====================================================
         SUPPORT
         ===================================================== */

      case 'support':
        return (
          <div className={styles.page}>
            <Support context={props.context} />
          </div>
        );

      /* =====================================================
         HOME
         ===================================================== */

      case 'home':
      default:
        return (
          <div className={styles.home}>

            <div className={styles.homeHeader}>

              <div className={styles.eyebrow}>
                CRM
              </div>

              <h1>
                Opportunities Tracker
              </h1>

              <p>
                Manage your sales opportunities from Sales
                through Tendering to Execution.
              </p>

            </div>

            <div className={styles.dashboardGrid}>

              {/* SALES */}

              <div
                className={styles.dashboardCard}
                onClick={() => setCurrentPage('sales')}
                role="button"
                tabIndex={0}
                onKeyDown={(event) => {
                  if (
                    event.key === 'Enter' ||
                    event.key === ' '
                  ) {
                    setCurrentPage('sales');
                  }
                }}
              >

                <div className={styles.cardTop}>

                  <div className={styles.cardIcon}>
                    S
                  </div>

                  <span className={styles.cardArrow}>
                    →
                  </span>

                </div>

                <h2>
                  Sales
                </h2>

                <p>
                  View and manage your sales opportunities,
                  track active deals, and monitor the pipeline.
                </p>

                <div className={styles.cardFooter}>

                  <span className={styles.cardLink}>
                    Open Sales
                  </span>

                  <span className={styles.cardFooterArrow}>
                    →
                  </span>

                </div>

              </div>

              {/* TENDERING */}

              <div
                className={styles.dashboardCard}
                onClick={() => setCurrentPage('tendering')}
                role="button"
                tabIndex={0}
                onKeyDown={(event) => {
                  if (
                    event.key === 'Enter' ||
                    event.key === ' '
                  ) {
                    setCurrentPage('tendering');
                  }
                }}
              >

                <div className={styles.cardTop}>

                  <div className={styles.cardIcon}>
                    T
                  </div>

                  <span className={styles.cardArrow}>
                    →
                  </span>

                </div>

                <h2>
                  Tendering
                </h2>

                <p>
                  Manage opportunities that have moved
                  from Sales into the tendering process.
                </p>

                <div className={styles.cardFooter}>

                  <span className={styles.cardLink}>
                    Open Tendering
                  </span>

                  <span className={styles.cardFooterArrow}>
                    →
                  </span>

                </div>

              </div>

              {/* EXECUTION */}

              <div
                className={styles.dashboardCard}
                onClick={() => setCurrentPage('execution')}
                role="button"
                tabIndex={0}
                onKeyDown={(event) => {
                  if (
                    event.key === 'Enter' ||
                    event.key === ' '
                  ) {
                    setCurrentPage('execution');
                  }
                }}
              >

                <div className={styles.cardTop}>

                  <div className={styles.cardIcon}>
                    E
                  </div>

                  <span className={styles.cardArrow}>
                    →
                  </span>

                </div>

                <h2>
                  Execution
                </h2>

                <p>
                  Track opportunities that were successfully
                  won and are ready for execution.
                </p>

                <div className={styles.cardFooter}>

                  <span className={styles.cardLink}>
                    Open Execution
                  </span>

                  <span className={styles.cardFooterArrow}>
                    →
                  </span>

                </div>

              </div>

            </div>

          </div>
        );
    }
  };

  return (
    <div className={styles.crm}>

      {/* =====================================================
          TOP NAVIGATION
          ===================================================== */}

      <header className={styles.topNavigation}>

        {/* BRAND */}

        <div
          className={styles.brand}
          onClick={() => setCurrentPage('home')}
          role="button"
          tabIndex={0}
          onKeyDown={(event) => {
            if (
              event.key === 'Enter' ||
              event.key === ' '
            ) {
              setCurrentPage('home');
            }
          }}
        >

          <div className={styles.logoIcon}>
            CRM
          </div>

          <div className={styles.brandText}>

            <strong>
              CRM
            </strong>

            <span>
              Opportunities Tracker
            </span>

          </div>

        </div>

        {/* =================================================
            NAVIGATION BUTTONS
            ================================================= */}

        <nav className={styles.navigation}>

          {/* HOME */}

          <button
            className={
              currentPage === 'home'
                ? styles.activeNav
                : styles.navButton
            }
            onClick={() => setCurrentPage('home')}
          >
            Home
          </button>

          {/* DASHBOARD */}

          <button
            className={
              currentPage === 'dashboard'
                ? styles.activeNav
                : styles.navButton
            }
            onClick={() => setCurrentPage('dashboard')}
          >
            Dashboard
          </button>

          {/* SALES */}

          <button
            className={
              currentPage === 'sales'
                ? styles.activeNav
                : styles.navButton
            }
            onClick={() => setCurrentPage('sales')}
          >
            Sales
          </button>

          {/* TENDERING */}

          <button
            className={
              currentPage === 'tendering'
                ? styles.activeNav
                : styles.navButton
            }
            onClick={() => setCurrentPage('tendering')}
          >
            Tendering
          </button>

          {/* EXECUTION */}

          <button
            className={
              currentPage === 'execution'
                ? styles.activeNav
                : styles.navButton
            }
            onClick={() => setCurrentPage('execution')}
          >
            Execution
          </button>

          {/* =================================================
              SUPPORT
              ================================================= */}

          <button
            className={
              currentPage === 'support'
                ? styles.activeNav
                : styles.navButton
            }
            onClick={() => setCurrentPage('support')}
          >
            Support
          </button>

        </nav>

      </header>

      {/* =====================================================
          MAIN CONTENT
          ===================================================== */}

      <main className={styles.mainContent}>

        {/* BACK BUTTON */}

        {currentPage !== 'home' && (

          <div className={styles.pageHeader}>

            <button
              className={styles.backButton}
              onClick={() => setCurrentPage('home')}
            >
              ← Back to Home
            </button>

          </div>

        )}

        {/* CURRENT PAGE */}

        {renderPage()}

      </main>

    </div>
  );
};

export default Crm;