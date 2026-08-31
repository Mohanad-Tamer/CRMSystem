import * as React from 'react';
import styles from './Dashboard.module.scss';
import type { IDashboardProps } from './IDashboardProps';

import { getSP } from '../../../pnpjsConfig';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

interface IOpportunity {
  Id: number;
  Title?: string;
  OpportunityNumber?: string;
  Category?: string;
  Client?: string;
  Stage?: string;
  TenderingStatus?: string;
  ExpectedValue?: number;
  SellingPercentage?: number;
  SalesMember?: {
    Title?: string;
  };
  ExpectedClose?: string;
  Team?: string;
  HandoverToTendering?: boolean;
  ProjectManager?: string;
}

interface IMonthData {
  month: string;
  value: number;
}

interface IStatusData {
  status: string;
  count: number;
}

interface IDomainData {
  domain: string;
  value: number;
  count: number;
}

const Dashboard: React.FC<IDashboardProps> = ({ context }) => {

  const [opportunities, setOpportunities] =
    React.useState<IOpportunity[]>([]);

  const [loading, setLoading] =
    React.useState<boolean>(true);

  const [error, setError] =
    React.useState<string>('');

  /*
   * =====================================================
   * LOAD DATA
   * =====================================================
   */

  const loadDashboard = async (): Promise<void> => {

    try {

      setLoading(true);
      setError('');

      const sp = getSP(context);

      const items: IOpportunity[] =
        await sp.web.lists
          .getByTitle('Opportunities')
          .items
          .select(
            'Id',
            'Title',
            'OpportunityNumber',
            'Category',
            'Client',
            'Stage',
            'TenderingStatus',
            'ExpectedValue',
            'SellingPercentage',
            'SalesMember/Title',
            'ExpectedClose',
            'Team',
            'HandoverToTendering',
            'ProjectManager'
          )
          .expand('SalesMember')();

      setOpportunities(items);

    } catch (err) {

      console.error(
        'Dashboard loading error:',
        err
      );

      setError(
        'Unable to load dashboard data from the Opportunities list.'
      );

    } finally {

      setLoading(false);

    }

  };

  React.useEffect(() => {

    void loadDashboard();

  }, []);

  /*
   * =====================================================
   * FORMATTERS
   * =====================================================
   */

  const formatCompactNumber = (
    value: number
  ): string => {

    const absoluteValue = Math.abs(value);

    if (absoluteValue >= 1000000000) {

      return (
        (value / 1000000000)
          .toFixed(1)
          .replace('.0', '') +
        'B'
      );

    }

    if (absoluteValue >= 1000000) {

      return (
        (value / 1000000)
          .toFixed(1)
          .replace('.0', '') +
        'M'
      );

    }

    if (absoluteValue >= 1000) {

      return (
        (value / 1000)
          .toFixed(1)
          .replace('.0', '') +
        'K'
      );

    }

    return String(
      Math.round(value)
    );

  };

  const formatCurrency = (
    value: number
  ): string => {

    return 'SAR ' + formatCompactNumber(value);

  };

  /*
   * =====================================================
   * BASIC METRICS
   * =====================================================
   */

  const totalOpportunities =
    opportunities.length;

  const pipelineValue =
    opportunities.reduce(
      (total, opportunity) =>
        total +
        (Number(opportunity.ExpectedValue) || 0),
      0
    );

  /*
   * =====================================================
   * WON
   *
   * TenderingStatus = Won
   * =====================================================
   */

  const wonOpportunities =
    opportunities.filter(
      opportunity =>
        (
          opportunity.TenderingStatus ||
          ''
        )
          .toLowerCase()
          .trim() === 'won'
    );

  const wonValue =
    wonOpportunities.reduce(
      (total, opportunity) =>
        total +
        (Number(opportunity.ExpectedValue) || 0),
      0
    );

  /*
   * =====================================================
   * READY FOR EXECUTION
   *
   * TenderingStatus = Won
   * AND
   * HandoverToTendering = Yes
   * =====================================================
   */

  const readyForExecution =
    opportunities.filter(
      opportunity =>
        (
          opportunity.TenderingStatus ||
          ''
        )
          .toLowerCase()
          .trim() === 'won' &&
        opportunity.HandoverToTendering === true
    );

  /*
   * =====================================================
   * LOST
   * =====================================================
   */

  const lostOpportunities =
    opportunities.filter(
      opportunity => {

        const stage =
          (
            opportunity.Stage ||
            ''
          )
            .toLowerCase()
            .trim();

        const tenderingStatus =
          (
            opportunity.TenderingStatus ||
            ''
          )
            .toLowerCase()
            .trim();

        return (
          stage === 'lost' ||
          tenderingStatus === 'lost'
        );

      }
    );

  /*
   * =====================================================
   * SELLING %
   * =====================================================
   */

  const sellingValues =
    opportunities
      .map(
        opportunity =>
          Number(
            opportunity.SellingPercentage
          )
      )
      .filter(
        value =>
          !isNaN(value)
      );

  const averageSelling =
    sellingValues.length > 0
      ? Math.round(
          sellingValues.reduce(
            (a, b) => a + b,
            0
          ) /
          sellingValues.length
        )
      : 0;

  /*
   * =====================================================
   * WIN RATE
   * =====================================================
   */

  const winRate =
    totalOpportunities > 0
      ? Math.round(
          (
            wonOpportunities.length /
            totalOpportunities
          ) *
          100
        )
      : 0;

  /*
   * =====================================================
   * SALES TO TENDERING
   * =====================================================
   */

  const handedOverCount =
    opportunities.filter(
      opportunity =>
        opportunity.HandoverToTendering === true
    ).length;

  const salesToTendering =
    totalOpportunities > 0
      ? Math.round(
          (
            handedOverCount /
            totalOpportunities
          ) *
          100
        )
      : 0;

  /*
   * =====================================================
   * STATUS DISTRIBUTION
   * =====================================================
   */

  const statusNames: string[] = [
    'Lead',
    'Qualification',
    'Proposal',
    'Negotiation',
    'Won',
    'Lost'
  ];

  const statusData: IStatusData[] =
    statusNames.map(
      status => {

        const count =
          opportunities.filter(
            opportunity =>
              (
                opportunity.Stage ||
                ''
              )
                .toLowerCase()
                .trim() ===
              status.toLowerCase()
          ).length;

        return {
          status,
          count
        };

      }
    );

  /*
   * =====================================================
   * STATUS COLORS
   *
   * Negotiation is intentionally PINK so it is clearly
   * different from Proposal.
   * =====================================================
   */

  const statusColors: {
    [key: string]: string;
  } = {
    Lead: '#3B82F6',
    Qualification: '#8B5CF6',
    Proposal: '#F59E0B',
    Negotiation: '#EC4899',
    Won: '#22C55E',
    Lost: '#EF4444'
  };

  /*
   * =====================================================
   * DONUT GRADIENT
   * =====================================================
   */

  const createDonutGradient = (): string => {

    if (totalOpportunities === 0) {
      return '#E5E7EB';
    }

    let currentPercentage = 0;

    const gradientParts: string[] = [];

    statusData.forEach(item => {

      const percentage =
        (
          item.count /
          totalOpportunities
        ) *
        100;

      if (percentage > 0) {

        const start =
          currentPercentage;

        const end =
          currentPercentage +
          percentage;

        gradientParts.push(
          `${statusColors[item.status]} ${start}% ${end}%`
        );

        currentPercentage =
          end;

      }

    });

    if (gradientParts.length === 0) {
      return '#E5E7EB';
    }

    return (
      'conic-gradient(' +
      gradientParts.join(', ') +
      ')'
    );

  };

  /*
   * =====================================================
   * DOMAIN DATA
   *
   * Uses Category column from SharePoint.
   * =====================================================
   */

  const domainNames: string[] = [];

  opportunities.forEach(
    opportunity => {

      const domain =
        opportunity.Category ||
        'Unassigned';

      if (
        domainNames.indexOf(domain) === -1
      ) {

        domainNames.push(domain);

      }

    }
  );

  const domainData: IDomainData[] =
    domainNames
      .map(
        domain => {

          const domainOpportunities =
            opportunities.filter(
              opportunity =>
                (
                  opportunity.Category ||
                  'Unassigned'
                ) === domain
            );

          const value =
            domainOpportunities.reduce(
              (total, opportunity) =>
                total +
                (
                  Number(
                    opportunity.ExpectedValue
                  ) || 0
                ),
              0
            );

          return {
            domain,
            value,
            count:
              domainOpportunities.length
          };

        }
      )
      .sort(
        (a, b) =>
          b.value - a.value
      );

  /*
   * =====================================================
   * PIPELINE TREND
   *
   * Last 6 months based on ExpectedClose.
   * Green dots have been REMOVED.
   * =====================================================
   */

  const createMonthKey = (
    date: Date
  ): string => {

    const month =
      date.getMonth() + 1;

    const monthString =
      month < 10
        ? '0' + month
        : String(month);

    return (
      date.getFullYear() +
      '-' +
      monthString
    );

  };

  const getMonthLabel = (
    date: Date
  ): string => {

    return date.toLocaleDateString(
      'en-US',
      {
        month: 'short'
      }
    );

  };

  const pipelineTrend: IMonthData[] =
    (() => {

      const result: IMonthData[] = [];

      const today =
        new Date();

      for (
        let i = 5;
        i >= 0;
        i--
      ) {

        const date =
          new Date(
            today.getFullYear(),
            today.getMonth() - i,
            1
          );

        const key =
          createMonthKey(date);

        const monthOpportunities =
          opportunities.filter(
            opportunity => {

              if (
                !opportunity.ExpectedClose
              ) {
                return false;
              }

              const opportunityDate =
                new Date(
                  opportunity.ExpectedClose
                );

              return (
                createMonthKey(
                  opportunityDate
                ) === key
              );

            }
          );

        const value =
          monthOpportunities.reduce(
            (total, opportunity) =>
              total +
              (
                Number(
                  opportunity.ExpectedValue
                ) || 0
              ),
            0
          );

        result.push({
          month:
            getMonthLabel(date),
          value
        });

      }

      return result;

    })();

  /*
   * =====================================================
   * PIPELINE CHART SCALE
   * =====================================================
   */

  const maxPipelineValue =
    Math.max(
      ...pipelineTrend.map(
        item =>
          item.value
      ),
      1
    );

  /*
   * =====================================================
   * RECENT OPPORTUNITIES
   * =====================================================
   */

  const recentOpportunities =
    [...opportunities]
      .sort(
        (a, b) =>
          b.Id - a.Id
      )
      .slice(0, 5);

  /*
   * =====================================================
   * LOADING
   * =====================================================
   */

  if (loading) {

    return (

      <div className={styles.dashboardPage}>

        <div className={styles.loading}>

          <div className={styles.loadingIcon}>
            ⏳
          </div>

          <h2>
            Loading dashboard...
          </h2>

          <p>
            Reading opportunities from SharePoint.
          </p>

        </div>

      </div>

    );

  }

  /*
   * =====================================================
   * ERROR
   * =====================================================
   */

  if (error) {

    return (

      <div className={styles.dashboardPage}>

        <div className={styles.errorState}>

          <div className={styles.errorIcon}>
            !
          </div>

          <h2>
            Unable to load dashboard
          </h2>

          <p>
            {error}
          </p>

          <button
            className={styles.refreshButton}
            onClick={() =>
              void loadDashboard()
            }
          >
            Try Again
          </button>

        </div>

      </div>

    );

  }

  /*
   * =====================================================
   * RENDER
   * =====================================================
   */

  return (

    <div className={styles.dashboardPage}>

      {/* =================================================
          HEADER
          ================================================= */}

      <div className={styles.dashboardHeader}>

        <div>

          <div className={styles.eyebrow}>
            CRM DASHBOARD
          </div>

          <h1>
            Executive Dashboard
          </h1>

          <p>
            Overview of opportunity pipeline
            and performance metrics.
          </p>

        </div>

        <button
          className={styles.refreshButton}
          onClick={() =>
            void loadDashboard()
          }
        >
          ↻ Refresh
        </button>

      </div>

      {/* =================================================
          KPI CARDS
          ================================================= */}

      <div className={styles.kpiGrid}>

        <div className={styles.kpiCard}>

          <div className={styles.kpiIcon}>
            📊
          </div>

          <div className={styles.kpiContent}>

            <span>
              TOTAL OPPORTUNITIES
            </span>

            <strong>
              {totalOpportunities}
            </strong>

          </div>

        </div>

        <div className={styles.kpiCard}>

          <div className={styles.kpiIcon}>
            💰
          </div>

          <div className={styles.kpiContent}>

            <span>
              PIPELINE VALUE
            </span>

            <strong>
              {formatCurrency(
                pipelineValue
              )}
            </strong>

          </div>

        </div>

        <div className={styles.kpiCard}>

          <div className={styles.kpiIcon}>
            🏆
          </div>

          <div className={styles.kpiContent}>

            <span>
              WON
            </span>

            <strong>
              {wonOpportunities.length}
            </strong>

          </div>

        </div>

        <div className={styles.kpiCard}>

          <div className={styles.kpiIcon}>
            🚀
          </div>

          <div className={styles.kpiContent}>

            <span>
              READY FOR EXECUTION
            </span>

            <strong>
              {readyForExecution.length}
            </strong>

          </div>

        </div>

      </div>

      {/* =================================================
          SECONDARY METRICS
          ================================================= */}

      <div className={styles.secondaryMetrics}>

        <div className={styles.metricCard}>

          <span>
            AVG. SELLING %
          </span>

          <strong>
            {averageSelling}%
          </strong>

        </div>

        <div className={styles.metricCard}>

          <span>
            SALES TO TENDERING
          </span>

          <strong>
            {salesToTendering}%
          </strong>

        </div>

        <div className={styles.metricCard}>

          <span>
            WIN RATE
          </span>

          <strong>
            {winRate}%
          </strong>

        </div>

        <div className={styles.metricCard}>

          <span>
            LOST
          </span>

          <strong>
            {lostOpportunities.length}
          </strong>

        </div>

      </div>

      {/* =================================================
          CHARTS
          ================================================= */}

      <div className={styles.chartGrid}>

        {/* =================================================
            STATUS DISTRIBUTION
            ================================================= */}

        <section className={styles.chartCard}>

          <div className={styles.cardHeader}>

            <div>

              <span className={styles.cardEyebrow}>
                PIPELINE
              </span>

              <h2>
                Status Distribution
              </h2>

              <p>
                Opportunities by current status
              </p>

            </div>

          </div>

          <div className={styles.statusChartArea}>

            {/* DONUT */}

            <div className={styles.donutWrapper}>

              <div
                className={styles.donut}
                style={{
                  background:
                    createDonutGradient()
                }}
              >

                <div className={styles.donutCenter}>

                  <strong>
                    {totalOpportunities}
                  </strong>

                  <span>
                    Total
                  </span>

                </div>

              </div>

            </div>

            {/* REFERENCE / LEGEND BELOW DONUT */}

            <div className={styles.legend}>

              {statusData.map(
                item => {

                  const percentage =
                    totalOpportunities > 0
                      ? Math.round(
                          (
                            item.count /
                            totalOpportunities
                          ) *
                          100
                        )
                      : 0;

                  return (

                    <div
                      key={item.status}
                      className={styles.legendItem}
                    >

                      <span
                        className={styles.legendDot}
                        style={{
                          background:
                            statusColors[item.status]
                        }}
                      />

                      <span>
                        {item.status}
                      </span>

                      <strong>
                        {item.count}
                      </strong>

                      <small>
                        {percentage}%
                      </small>

                    </div>

                  );

                }
              )}

            </div>

          </div>

        </section>

        {/* =================================================
            OPPORTUNITIES BY DOMAIN
            ================================================= */}

        <section className={styles.chartCard}>

          <div className={styles.cardHeader}>

            <div>

              <span className={styles.cardEyebrow}>
                DOMAIN
              </span>

              <h2>
                Opportunities by Domain
              </h2>

              <p>
                Distribution across business domains
              </p>

            </div>

          </div>

          <div className={styles.domainChart}>

            {domainData.length === 0 ? (

              <div className={styles.noData}>
                No domain data available.
              </div>

            ) : (

              domainData
                .slice(0, 6)
                .map(
                  item => {

                    const percentage =
                      pipelineValue > 0
                        ? Math.round(
                            (
                              item.value /
                              pipelineValue
                            ) *
                            100
                          )
                        : 0;

                    return (

                      <div
                        key={item.domain}
                        className={styles.domainRow}
                      >

                        <div
                          className={
                            styles.domainHeader
                          }
                        >

                          <span>
                            {item.domain}
                          </span>

                          <strong>
                            {formatCurrency(
                              item.value
                            )}
                          </strong>

                        </div>

                        <div
                          className={
                            styles.domainBar
                          }
                        >

                          <div
                            className={
                              styles.domainFill
                            }
                            style={{
                              width:
                                `${percentage}%`
                            }}
                          />

                        </div>

                        <div
                          className={
                            styles.domainFooter
                          }
                        >

                          <span>
                            {item.count}
                            {' '}
                            opportunities
                          </span>

                          <span>
                            {percentage}%
                          </span>

                        </div>

                      </div>

                    );

                  }
                )

            )}

          </div>

        </section>

      </div>

      {/* =================================================
          PIPELINE TREND
          ================================================= */}

      <section className={styles.chartCard}>

        <div className={styles.cardHeader}>

          <div>

            <span className={styles.cardEyebrow}>
              TREND
            </span>

            <h2>
              Pipeline Trend
            </h2>

            <p>
              Monthly pipeline value
            </p>

          </div>

          <div className={styles.chartLegend}>

            <span>

              <i
                className={
                  styles.pipelineLegend
                }
              />

              Pipeline

            </span>

          </div>

        </div>

        <div className={styles.pipelineChart}>

          {/* Y AXIS */}

          <div className={styles.chartYAxis}>

            <span>
              {formatCompactNumber(
                maxPipelineValue
              )}
            </span>

            <span>
              {formatCompactNumber(
                maxPipelineValue * 0.75
              )}
            </span>

            <span>
              {formatCompactNumber(
                maxPipelineValue * 0.5
              )}
            </span>

            <span>
              {formatCompactNumber(
                maxPipelineValue * 0.25
              )}
            </span>

            <span>
              0
            </span>

          </div>

          {/* CHART */}

          <div className={styles.chartPlot}>

            <div className={styles.chartGridLines}>

              <span />
              <span />
              <span />
              <span />
              <span />

            </div>

            <div className={styles.barChart}>

              {pipelineTrend.map(
                item => {

                  const height =
                    Math.max(
                      (
                        item.value /
                        maxPipelineValue
                      ) * 100,
                      item.value > 0
                        ? 3
                        : 0
                    );

                  return (

                    <div
                      key={item.month}
                      className={
                        styles.pipelineColumn
                      }
                    >

                      <div
                        className={
                          styles.pipelineTooltip
                        }
                      >
                        {formatCurrency(
                          item.value
                        )}
                      </div>

                      <div
                        className={
                          styles.pipelineBar
                        }
                        style={{
                          height:
                            `${height}%`
                        }}
                      />

                      <span
                        className={
                          styles.pipelineMonth
                        }
                      >
                        {item.month}
                      </span>

                    </div>

                  );

                }
              )}

            </div>

          </div>

        </div>

      </section>

      {/* =================================================
          SUMMARY
          ================================================= */}

      <section className={styles.summaryCard}>

        <div className={styles.summaryItem}>

          <span>
            WON VALUE
          </span>

          <strong>
            {formatCurrency(
              wonValue
            )}
          </strong>

        </div>

        <div className={styles.summaryDivider} />

        <div className={styles.summaryItem}>

          <span>
            ACTIVE PIPELINE
          </span>

          <strong>
            {formatCurrency(
              Math.max(
                pipelineValue - wonValue,
                0
              )
            )}
          </strong>

        </div>

        <div className={styles.summaryDivider} />

        <div className={styles.summaryItem}>

          <span>
            READY FOR EXECUTION
          </span>

          <strong>
            {readyForExecution.length}
          </strong>

        </div>

      </section>

      {/* =================================================
          RECENT OPPORTUNITIES
          ================================================= */}

      <section className={styles.dashboardCard}>

        <div className={styles.cardHeader}>

          <div>

            <span className={styles.cardEyebrow}>
              ACTIVITY
            </span>

            <h2>
              Recent Opportunities
            </h2>

            <p>
              Latest activity in the pipeline
            </p>

          </div>

        </div>

        {recentOpportunities.length === 0 ? (

          <div className={styles.noData}>
            No opportunities found.
          </div>

        ) : (

          <div className={styles.tableWrapper}>

            <table
              className={
                styles.opportunityTable
              }
            >

              <thead>

                <tr>

                  <th>
                    OPPORTUNITY
                  </th>

                  <th>
                    CLIENT
                  </th>

                  <th>
                    VALUE
                  </th>

                  <th>
                    CATEGORY
                  </th>

                  <th>
                    STATUS
                  </th>

                </tr>

              </thead>

              <tbody>

                {recentOpportunities.map(
                  opportunity => (

                    <tr
                      key={
                        opportunity.Id
                      }
                    >

                      <td>

                        <strong>
                          {
                            opportunity.OpportunityNumber ||
                            opportunity.Title ||
                            '-'
                          }
                        </strong>

                        {opportunity.Title && (

                          <small>
                            {opportunity.Title}
                          </small>

                        )}

                      </td>

                      <td>
                        {
                          opportunity.Client ||
                          '-'
                        }
                      </td>

                      <td>
                        {formatCurrency(
                          Number(
                            opportunity.ExpectedValue
                          ) || 0
                        )}
                      </td>

                      <td>
                        {
                          opportunity.Category ||
                          '-'
                        }
                      </td>

                      <td>

                        <span
                          className={
                            opportunity.TenderingStatus &&
                            opportunity.TenderingStatus
                              .toLowerCase()
                              .trim() === 'won'
                              ? styles.wonBadge
                              : styles.stageBadge
                          }
                        >
                          {
                            opportunity.TenderingStatus ||
                            opportunity.Stage ||
                            '-'
                          }
                        </span>

                      </td>

                    </tr>

                  )
                )}

              </tbody>

            </table>

          </div>

        )}

      </section>

    </div>

  );

};

export default Dashboard;