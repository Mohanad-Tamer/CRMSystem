import * as React from 'react';
import { useEffect, useState } from 'react';

import styles from './Tendering.module.scss';

import type {
  ITenderingProps,
  ITenderRecord
} from './ITenderingProps';

import { getSP } from '../../../pnpjsConfig';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

const Tendering: React.FC<ITenderingProps> = ({
  context
}) => {

  const [records, setRecords] =
    useState<ITenderRecord[]>([]);

  const [searchTerm, setSearchTerm] =
    useState<string>('');

  const [statusFilter, setStatusFilter] =
    useState<string>('All');

  const [loading, setLoading] =
    useState<boolean>(true);

  const [error, setError] =
    useState<string>('');

  /*
   * LOAD DATA FROM OPPORTUNITIES LIST
   */

  const loadRecords = async (): Promise<void> => {

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
          'Category',
          'Client',
          'Stage',
          'TenderingStatus',
          'ExpectedClose',
          'Description',
          'Team',
          'TenderingManager',
          'TenderingEngineer',
          'HandoverToTendering',
          'Modified'
        )
        .filter(
          'HandoverToTendering eq 1'
        )
        .orderBy(
          'ExpectedClose',
          true
        )();

      const tenderRecords: ITenderRecord[] =
        items.map((item: any) => ({
          Id: item.Id,

          title:
            item.Title || '',

          opportunityNumber:
            item.OpportunityNumber || '',

          type:
            item.Category || '',

          status:
            item.TenderingStatus ||
            item.Stage ||
            'In Progress',

          customer:
            item.Client || '',

          deadline:
            item.ExpectedClose || '',

          manager:
            item.TenderingManager || '',

          engineer:
            item.TenderingEngineer || '',

          comment:
            item.Description || '',

          domain:
            item.Team || '',

          lastUpdated:
            item.Modified || '',

          handoverToTendering:
            item.HandoverToTendering === true
        }));

      setRecords(tenderRecords);

    } catch (err) {

      console.error(
        'Error loading tendering records:',
        err
      );

      setError(
        'Unable to load tendering records from the Opportunities list.'
      );

    } finally {

      setLoading(false);

    }

  };

  /*
   * LOAD WHEN WEB PART OPENS
   */

  useEffect(() => {

    void loadRecords();

  }, []);

  /*
   * MARK WON
   */

  const markWon = async (
    id: number
  ): Promise<void> => {

    try {

      const sp = getSP(context);

      await sp.web.lists
        .getByTitle('Opportunities')
        .items
        .getById(id)
        .update({
          TenderingStatus: 'Won'
        });

      setRecords(previous =>
        previous.map(record =>
          record.Id === id
            ? {
                ...record,
                status: 'Won'
              }
            : record
        )
      );

    } catch (err) {

      console.error(
        'Error marking record as Won:',
        err
      );

      setError(
        'Unable to update the tender status.'
      );

    }

  };

  /*
   * MARK LOST
   */

  const markLost = async (
    id: number
  ): Promise<void> => {

    try {

      const sp = getSP(context);

      await sp.web.lists
        .getByTitle('Opportunities')
        .items
        .getById(id)
        .update({
          TenderingStatus: 'Lost'
        });

      setRecords(previous =>
        previous.map(record =>
          record.Id === id
            ? {
                ...record,
                status: 'Lost'
              }
            : record
        )
      );

    } catch (err) {

      console.error(
        'Error marking record as Lost:',
        err
      );

      setError(
        'Unable to update the tender status.'
      );

    }

  };

  /*
   * SEARCH + FILTER
   */

  const filteredRecords =
    records.filter(record => {

      const search =
        searchTerm.toLowerCase().trim();

      const matchesSearch =
        !search ||
        record.title
          .toLowerCase()
          .includes(search) ||
        record.opportunityNumber
          .toLowerCase()
          .includes(search) ||
        record.customer
          .toLowerCase()
          .includes(search) ||
        record.manager
          .toLowerCase()
          .includes(search) ||
        record.engineer
          .toLowerCase()
          .includes(search);

      const matchesStatus =
        statusFilter === 'All' ||
        record.status === statusFilter;

      return (
        matchesSearch &&
        matchesStatus
      );

    });

  /*
   * FORMAT DATE
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
   * SUMMARY COUNTS
   */

  const totalRecords =
    records.length;

  const inProgressCount =
    records.filter(
      record =>
        record.status === 'In Progress'
    ).length;

  const wonCount =
    records.filter(
      record =>
        record.status === 'Won'
    ).length;

  const lostCount =
    records.filter(
      record =>
        record.status === 'Lost'
    ).length;

  /*
   * RENDER
   */

  return (

    <div className={styles.tenderingPage}>

      {/* SUMMARY */}

      <div className={styles.summaryGrid}>

        <div className={styles.summaryCard}>

          <div className={styles.summaryIcon}>
            📋
          </div>

          <div>
            <span>
              Total Records
            </span>

            <strong>
              {totalRecords}
            </strong>
          </div>

        </div>

        <div className={styles.summaryCard}>

          <div className={styles.summaryIcon}>
            ⏳
          </div>

          <div>
            <span>
              In Progress
            </span>

            <strong>
              {inProgressCount}
            </strong>
          </div>

        </div>

        <div className={styles.summaryCard}>

          <div className={styles.summaryIcon}>
            ✓
          </div>

          <div>
            <span>
              Won
            </span>

            <strong>
              {wonCount}
            </strong>
          </div>

        </div>

        <div className={styles.summaryCard}>

          <div className={styles.summaryIcon}>
            ✕
          </div>

          <div>
            <span>
              Lost
            </span>

            <strong>
              {lostCount}
            </strong>
          </div>

        </div>

      </div>

      {/* TOOLBAR */}

      <div className={styles.toolbar}>

        <div className={styles.searchBox}>

          <span>
            🔍
          </span>

          <input
            type="text"
            placeholder="Search tendering records..."
            value={searchTerm}
            onChange={event =>
              setSearchTerm(
                event.target.value
              )
            }
          />

        </div>

        <select
          className={styles.statusFilter}
          value={statusFilter}
          onChange={event =>
            setStatusFilter(
              event.target.value
            )
          }
        >

          <option value="All">
            All Statuses
          </option>

          <option value="In Progress">
            In Progress
          </option>

          <option value="Won">
            Won
          </option>

          <option value="Lost">
            Lost
          </option>

          <option value="Pending">
            Pending
          </option>

        </select>

      </div>

      {/* LOADING */}

      {loading && (

        <div className={styles.noRecords}>
          Loading tendering records...
        </div>

      )}

      {/* ERROR */}

      {error && (

        <div className={styles.noRecords}>
          {error}
        </div>

      )}

      {/* EMPTY */}

      {!loading &&
        !error &&
        filteredRecords.length === 0 && (

          <div className={styles.noRecords}>
            No tendering records found.
          </div>

        )}

      {/* RECORDS */}

      {!loading &&
        filteredRecords.map(record => {

          const isWon =
            record.status === 'Won';

          const isLost =
            record.status === 'Lost';

          return (

            <div
              key={record.Id}
              className={
                isWon
                  ? `${styles.tenderCard} ${styles.wonCard}`
                  : isLost
                    ? `${styles.tenderCard} ${styles.lostCard}`
                    : styles.tenderCard
              }
            >

              {/* HEADER */}

              <div className={styles.cardTop}>

                <div className={styles.recordInfo}>

                  <div className={styles.recordId}>
                    {record.opportunityNumber ||
                      `#${record.Id}`}
                  </div>

                  <span className={styles.typeBadge}>
                    {record.type || '-'}
                  </span>

                  <span className={styles.statusBadge}>
                    {record.status}
                  </span>

                </div>

                <div className={styles.domain}>
                  {record.domain || '-'}
                </div>

              </div>

              {/* TITLE */}

              <h2>
                {record.title || '-'}
              </h2>

              <div className={styles.customer}>
                {record.customer || '-'}
              </div>

              {/* DETAILS */}

              <div className={styles.detailsGrid}>

                <div className={styles.detailItem}>

                  <span className={styles.detailLabel}>
                    Deadline
                  </span>

                  <strong>
                    {formatDate(
                      record.deadline
                    )}
                  </strong>

                </div>

                <div className={styles.detailItem}>

                  <span className={styles.detailLabel}>
                    Tendering Manager
                  </span>

                  <strong>
                    {record.manager || '-'}
                  </strong>

                </div>

                <div className={styles.detailItem}>

                  <span className={styles.detailLabel}>
                    Tendering Engineer
                  </span>

                  <strong>
                    {record.engineer || '-'}
                  </strong>

                </div>

              </div>

              {/* DESCRIPTION */}

              <div className={styles.comment}>

                <span>
                  💬
                </span>

                {record.comment ||
                  'No comment provided.'}

              </div>

              {/* FOOTER */}

              <div className={styles.cardFooter}>

                <div>

                  <span>
                    Last Updated
                  </span>

                  <strong>
                    {formatDate(
                      record.lastUpdated
                    )}
                  </strong>

                </div>

                <div className={styles.actions}>

                  {!isWon &&
                    !isLost && (

                      <>
                        <button
                          className={
                            styles.wonButton
                          }
                          onClick={() =>
                            void markWon(
                              record.Id
                            )
                          }
                        >
                          ✓ Mark Won
                        </button>

                        <button
                          className={
                            styles.lostButton
                          }
                          onClick={() =>
                            void markLost(
                              record.Id
                            )
                          }
                        >
                          ✕ Mark Lost
                        </button>
                      </>

                    )}

                  {isWon && (

                    <span
                      className={
                        styles.handedOverText
                      }
                    >
                      ✓ Won
                    </span>

                  )}

                  {isLost && (

                    <span>
                      ✕ Lost
                    </span>

                  )}

                </div>

              </div>

            </div>

          );

        })}

    </div>

  );

};

export default Tendering;