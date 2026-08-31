import * as React from 'react';
import { useEffect, useState } from 'react';

import styles from './Opportunities.module.scss';
import { IOpportunitiesProps } from './IOpportunitiesProps';
import { getSP } from '../../../pnpjsConfig';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

interface IOpportunity {
  Id: number;
  Title: string;
  OpportunityNumber: string;
  Revision: number;
  Category: string;
  Client: string;
  Stage: string;
  ExpectedValue: number;
  SellingPercentage: number;
  SalesMember?: {
    Title: string;
  };
  ExpectedClose: string;
  Description: string;
  HandoverToTendering: boolean;
  Team: string;
  Modified: string;
}

const Opportunities: React.FC<IOpportunitiesProps> = ({
  context
}) => {

  const [opportunities, setOpportunities] =
    useState<IOpportunity[]>([]);

  const [searchText, setSearchText] =
    useState<string>('');

  const [stageFilter, setStageFilter] =
    useState<string>('All');

  const [categoryFilter, setCategoryFilter] =
    useState<string>('All');

  const [loading, setLoading] =
    useState<boolean>(true);

  const [error, setError] =
    useState<string>('');

  const [handoverId, setHandoverId] =
    useState<number | null>(null);

  /*
   * LOAD OPPORTUNITIES
   */

  const loadOpportunities = async (): Promise<void> => {

    try {

      setLoading(true);
      setError('');

      const sp = getSP(context);

      const items = await sp.web.lists
        .getByTitle('Opportunities')
        .items
        .select(
          'Id',
          'Title',
          'OpportunityNumber',
          'Revision',
          'Category',
          'Client',
          'Stage',
          'ExpectedValue',
          'SellingPercentage',
          'SalesMember/Title',
          'ExpectedClose',
          'Description',
          'HandoverToTendering',
          'Team',
          'Modified'
        )
        .expand('SalesMember')
        .orderBy('Modified', false)();

      setOpportunities(items as IOpportunity[]);

    } catch (err) {

      console.error(
        'Error loading opportunities:',
        err
      );

      setError(
        'Unable to load opportunities. Please check your SharePoint list and column names.'
      );

    } finally {

      setLoading(false);

    }

  };

  /*
   * HANDOVER TO TENDERING
   */

  const handoverToTendering = async (
    opportunity: IOpportunity
  ): Promise<void> => {

    try {

      setHandoverId(opportunity.Id);
      setError('');

      const sp = getSP(context);

      await sp.web.lists
        .getByTitle('Opportunities')
        .items
        .getById(opportunity.Id)
        .update({
          HandoverToTendering: true,
          Stage: 'Handed Over to Tendering'
        });

      setOpportunities(previous =>
        previous.map(item => {

          if (item.Id === opportunity.Id) {

            return {
              ...item,
              HandoverToTendering: true,
              Stage: 'Handed Over to Tendering'
            };

          }

          return item;

        })
      );

    } catch (err) {

      console.error(
        'Error handing over opportunity:',
        err
      );

      setError(
        'Unable to hand over this opportunity to Tendering.'
      );

    } finally {

      setHandoverId(null);

    }

  };

  /*
   * LOAD DATA
   */

  useEffect(() => {

    void loadOpportunities();

  }, []);

  /*
   * FILTER
   */

  const filteredOpportunities =
    opportunities.filter(opportunity => {

      const search =
        searchText.toLowerCase();

      const matchesSearch =
        !search ||
        opportunity.Title?.toLowerCase().includes(search) ||
        opportunity.Client?.toLowerCase().includes(search) ||
        opportunity.OpportunityNumber?.toLowerCase().includes(search) ||
        opportunity.Description?.toLowerCase().includes(search);

      const matchesStage =
        stageFilter === 'All' ||
        opportunity.Stage === stageFilter;

      const matchesCategory =
        categoryFilter === 'All' ||
        opportunity.Category === categoryFilter;

      return (
        matchesSearch &&
        matchesStage &&
        matchesCategory
      );

    });

  /*
   * CURRENCY
   */

  const formatCurrency = (
    value: number
  ): string => {

    if (!value) {
      return '$0';
    }

    return new Intl.NumberFormat(
      'en-US',
      {
        style: 'currency',
        currency: 'USD',
        maximumFractionDigits: 0
      }
    ).format(value);

  };

  /*
   * DATE
   */

  const formatDate = (
    date: string
  ): string => {

    if (!date) {
      return '-';
    }

    return new Date(date)
      .toLocaleDateString(
        'en-US',
        {
          month: 'short',
          day: 'numeric',
          year: 'numeric'
        }
      );

  };

  /*
   * STAGE COLOR
   */

  const getStageClass = (
    stage: string
  ): string => {

    switch (stage) {

      case 'Proposal':
        return styles.stageProposal;

      case 'Negotiation':
        return styles.stageNegotiation;

      case 'Qualification':
        return styles.stageQualification;

      case 'New Lead':
        return styles.stageNewLead;

      case 'Handed Over to Tendering':
        return styles.stageTendering;

      default:
        return styles.stageDefault;

    }

  };

  /*
   * STAGES
   */

  const stages = Array.from(
    new Set(
      opportunities
        .map(item => item.Stage)
        .filter(Boolean)
    )
  );

  /*
   * CATEGORIES
   */

  const categories = Array.from(
    new Set(
      opportunities
        .map(item => item.Category)
        .filter(Boolean)
    )
  );

  return (

    <div className={styles.salesPage}>

      <main className={styles.main}>

        

        {/* SEARCH AND FILTERS */}

        <section className={styles.toolbar}>

          <div className={styles.searchContainer}>

            <span className={styles.searchIcon}>
              🔍
            </span>

            <input
              type="text"
              placeholder="Search opportunities..."
              value={searchText}
              onChange={event =>
                setSearchText(
                  event.target.value
                )
              }
            />

          </div>

          <select
            className={styles.filter}
            value={stageFilter}
            onChange={event =>
              setStageFilter(
                event.target.value
              )
            }
          >

            <option value="All">
              All Stages
            </option>

            {stages.map(stage => (

              <option
                key={stage}
                value={stage}
              >
                {stage}
              </option>

            ))}

          </select>

          <select
            className={styles.filter}
            value={categoryFilter}
            onChange={event =>
              setCategoryFilter(
                event.target.value
              )
            }
          >

            <option value="All">
              All Categories
            </option>

            {categories.map(category => (

              <option
                key={category}
                value={category}
              >
                {category}
              </option>

            ))}

          </select>

        </section>

        {/* RESULT COUNT */}

        <div className={styles.resultCount}>

          Showing{' '}

          <strong>
            {filteredOpportunities.length}
          </strong>

          {' '}of{' '}

          <strong>
            {opportunities.length}
          </strong>

          {' '}opportunities

        </div>

        {/* LOADING */}

        {loading && (

          <div className={styles.message}>
            Loading opportunities...
          </div>

        )}

        {/* ERROR */}

        {error && (

          <div className={styles.error}>
            {error}
          </div>

        )}

        {/* EMPTY */}

        {!loading &&
          !error &&
          filteredOpportunities.length === 0 && (

            <div className={styles.empty}>
              No opportunities found.
            </div>

          )}

        {/* OPPORTUNITY CARDS */}

        <div className={styles.cards}>

          {filteredOpportunities.map(
            opportunity => {

              const handedOver =
                opportunity.HandoverToTendering;

              return (

                <article
                  key={opportunity.Id}
                  className={
                    handedOver
                      ? `${styles.card} ${styles.cardHandedOver}`
                      : styles.card
                  }
                >

                  {/* CARD HEADER */}

                  <div className={styles.cardHeader}>

                    <div>

                      <div
                        className={
                          styles.opportunityNumber
                        }
                      >

                        {opportunity.OpportunityNumber ||
                          'No Number'}

                        <span
                          className={
                            styles.revision
                          }
                        >

                          Rev.{' '}

                          {opportunity.Revision ||
                            0}

                        </span>

                      </div>

                      <h2>
                        {opportunity.Title}
                      </h2>

                    </div>

                    <span
                      className={`${styles.stage} ${getStageClass(
                        opportunity.Stage
                      )}`}
                    >
                      {opportunity.Stage}
                    </span>

                  </div>

                  {/* CLIENT / CATEGORY */}

                  <div className={styles.cardInfo}>

                    <div
                      className={
                        styles.infoBlock
                      }
                    >

                      <span
                        className={
                          styles.label
                        }
                      >
                        CLIENT
                      </span>

                      <span
                        className={
                          styles.value
                        }
                      >
                        {opportunity.Client ||
                          '-'}
                      </span>

                    </div>

                    <div
                      className={
                        styles.infoBlock
                      }
                    >

                      <span
                        className={
                          styles.label
                        }
                      >
                        CATEGORY
                      </span>

                      <span
                        className={
                          styles.value
                        }
                      >
                        {opportunity.Category ||
                          '-'}
                      </span>

                    </div>

                  </div>

                  {/* METRICS */}

                  <div className={styles.metrics}>

                    <div className={styles.metric}>

                      <span
                        className={
                          styles.label
                        }
                      >
                        EXPECTED VALUE
                      </span>

                      <span
                        className={
                          styles.metricValue
                        }
                      >
                        {formatCurrency(
                          opportunity.ExpectedValue
                        )}
                      </span>

                    </div>

                    <div className={styles.metric}>

                      <span
                        className={
                          styles.label
                        }
                      >
                        SELLING %
                      </span>

                      <span
                        className={
                          styles.metricValue
                        }
                      >
                        {opportunity.SellingPercentage ||
                          0}%
                      </span>

                    </div>

                  </div>

                  {/* SALES INFORMATION */}

                  <div className={styles.cardInfo}>

                    <div
                      className={
                        styles.infoBlock
                      }
                    >

                      <span
                        className={
                          styles.label
                        }
                      >
                        SALES MEMBER
                      </span>

                      <span
                        className={
                          styles.value
                        }
                      >
                        {opportunity.SalesMember?.Title ||
                          '-'}
                      </span>

                    </div>

                    <div
                      className={
                        styles.infoBlock
                      }
                    >

                      <span
                        className={
                          styles.label
                        }
                      >
                        EXPECTED CLOSE
                      </span>

                      <span
                        className={
                          styles.value
                        }
                      >
                        {formatDate(
                          opportunity.ExpectedClose
                        )}
                      </span>

                    </div>

                  </div>

                  {/* DESCRIPTION */}

                  <div
                    className={
                      styles.description
                    }
                  >

                    <span
                      className={
                        styles.label
                      }
                    >
                      DESCRIPTION
                    </span>

                    <p>
                      {opportunity.Description ||
                        'No description provided.'}
                    </p>

                  </div>

                  {/* FOOTER */}

                  <div
                    className={
                      styles.cardFooter
                    }
                  >

                    <div
                      className={
                        styles.footerItem
                      }
                    >

                      <span
                        className={
                          styles.label
                        }
                      >
                        HANDOVER
                      </span>

                      {!handedOver ? (

                        <button
                          className={
                            styles.handoverButton
                          }
                          disabled={
                            handoverId ===
                            opportunity.Id
                          }
                          onClick={() =>
                            void handoverToTendering(
                              opportunity
                            )
                          }
                        >

                          {handoverId ===
                          opportunity.Id
                            ? 'Handing over...'
                            : 'Handover to Tendering'}

                        </button>

                      ) : (

                        <span
                          className={
                            styles.handedOverText
                          }
                        >
                          ✓ Handed Over
                        </span>

                      )}

                    </div>

                    <div
                      className={
                        styles.footerItem
                      }
                    >

                      <span
                        className={
                          styles.label
                        }
                      >
                        TEAM
                      </span>

                      <span>
                        {opportunity.Team ||
                          '-'}
                      </span>

                    </div>

                    <div
                      className={
                        styles.footerItem
                      }
                    >

                      <span
                        className={
                          styles.label
                        }
                      >
                        UPDATED
                      </span>

                      <span>
                        {formatDate(
                          opportunity.Modified
                        )}
                      </span>

                    </div>

                  </div>

                </article>

              );

            }
          )}

        </div>

      </main>

    </div>

  );

};

export default Opportunities;