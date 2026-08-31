import * as React from 'react';
import styles from './Execution.module.scss';
import type { IExecutionProps } from './IExecutionProps';

import { getSP } from '../../../pnpjsConfig';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

interface IExecutionProject {
  sharePointId: number;

  id: string;
  title: string;
  customer: string;
  domain: string;
  manager: string;
  engineer: string;
  deadline: string;

  projectManager: string;

  checklist: {
    signedContract: boolean;
    finalTechnicalProposal: boolean;
    finalCommercialProposal: boolean;
    scopeOfWork: boolean;
    clientRequirements: boolean;
    riskRegister: boolean;
    projectTimeline: boolean;
    stakeholderContacts: boolean;
  };
}

interface ISharePointOpportunity {
  Id: number;

  Title?: string;
  OpportunityNumber?: string;
  Category?: string;
  Client?: string;
  Stage?: string;
  ExpectedClose?: string;
  Team?: string;
  TenderingManager?: string;
  TenderingEngineer?: string;
  HandoverToTendering?: boolean;

  ProjectManager?: string;

  ExecutionChecklist?: unknown;
}

interface IChecklistItem {
  key: keyof IExecutionProject['checklist'];
  label: string;
}

/*
 * Project Manager choices
 */

const projectManagers: string[] = [
  'Ahmed Hassan',
  'Mohamed Ali',
  'Omar Nasser',
  'Mostafa Adel'
];

/*
 * Execution checklist
 *
 * These names MUST exactly match
 * the choices in SharePoint.
 */

const checklistItems: IChecklistItem[] = [
  {
    key: 'signedContract',
    label: 'Signed Contract / PO'
  },
  {
    key: 'finalTechnicalProposal',
    label: 'Final Technical Proposal'
  },
  {
    key: 'finalCommercialProposal',
    label: 'Final Commercial Proposal'
  },
  {
    key: 'scopeOfWork',
    label: 'Scope of Work'
  },
  {
    key: 'clientRequirements',
    label: 'Client Requirements'
  },
  {
    key: 'riskRegister',
    label: 'Risk Register'
  },
  {
    key: 'projectTimeline',
    label: 'Project Timeline'
  },
  {
    key: 'stakeholderContacts',
    label: 'Stakeholder Contacts'
  }
];

/*
 * Convert SharePoint multi-choice value
 * into a normal string array.
 */

const getChecklistValues = (
  value: unknown
): string[] => {

  if (!value) {
    return [];
  }

  /*
   * Example:
   * ["Signed Contract / PO", "Scope of Work"]
   */

  if (Array.isArray(value)) {

    return value.filter(
      item =>
        typeof item === 'string'
    ) as string[];

  }

  /*
   * Example:
   * {
   *   results: [
   *     "Signed Contract / PO"
   *   ]
   * }
   */

  if (
    typeof value === 'object' &&
    value !== null
  ) {

    const objectValue =
      value as {
        results?: unknown;
      };

    if (
      Array.isArray(
        objectValue.results
      )
    ) {

      return objectValue.results.filter(
        item =>
          typeof item === 'string'
      ) as string[];

    }

  }

  return [];
};

/*
 * Convert SharePoint checklist
 * into React boolean values.
 */

const createChecklist = (
  values: string[]
): IExecutionProject['checklist'] => {

  return {

    signedContract:
      values.indexOf(
        'Signed Contract / PO'
      ) !== -1,

    finalTechnicalProposal:
      values.indexOf(
        'Final Technical Proposal'
      ) !== -1,

    finalCommercialProposal:
      values.indexOf(
        'Final Commercial Proposal'
      ) !== -1,

    scopeOfWork:
      values.indexOf(
        'Scope of Work'
      ) !== -1,

    clientRequirements:
      values.indexOf(
        'Client Requirements'
      ) !== -1,

    riskRegister:
      values.indexOf(
        'Risk Register'
      ) !== -1,

    projectTimeline:
      values.indexOf(
        'Project Timeline'
      ) !== -1,

    stakeholderContacts:
      values.indexOf(
        'Stakeholder Contacts'
      ) !== -1

  };

};

/*
 * Convert React checklist into
 * SharePoint multi-choice values.
 */

const checklistToSharePoint = (
  checklist: IExecutionProject['checklist']
): string[] => {

  const values: string[] = [];

  if (
    checklist.signedContract
  ) {
    values.push(
      'Signed Contract / PO'
    );
  }

  if (
    checklist.finalTechnicalProposal
  ) {
    values.push(
      'Final Technical Proposal'
    );
  }

  if (
    checklist.finalCommercialProposal
  ) {
    values.push(
      'Final Commercial Proposal'
    );
  }

  if (
    checklist.scopeOfWork
  ) {
    values.push(
      'Scope of Work'
    );
  }

  if (
    checklist.clientRequirements
  ) {
    values.push(
      'Client Requirements'
    );
  }

  if (
    checklist.riskRegister
  ) {
    values.push(
      'Risk Register'
    );
  }

  if (
    checklist.projectTimeline
  ) {
    values.push(
      'Project Timeline'
    );
  }

  if (
    checklist.stakeholderContacts
  ) {
    values.push(
      'Stakeholder Contacts'
    );
  }

  return values;
};

const Execution: React.FC<IExecutionProps> = ({
  context
}) => {

  const [projects, setProjects] =
    React.useState<IExecutionProject[]>([]);

  const [loading, setLoading] =
    React.useState<boolean>(true);

  const [error, setError] =
    React.useState<string>('');

  const [saving, setSaving] =
    React.useState<boolean>(false);

  const [editingProjectId, setEditingProjectId] =
    React.useState<string | null>(null);

  const [searchTerm, setSearchTerm] =
    React.useState<string>('');

  const [sortBy, setSortBy] =
    React.useState<string>('default');

  /*
   * LOAD WON OPPORTUNITIES
   */

  const loadProjects = async (): Promise<void> => {

    try {

      setLoading(true);
      setError('');

      const sp = getSP(context);

      const items: ISharePointOpportunity[] =
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
            'ExpectedClose',
            'Team',
            'TenderingManager',
            'TenderingEngineer',
            'HandoverToTendering',
            'ProjectManager',
            'ExecutionChecklist'
          )
          .filter(
              "TenderingStatus eq 'Won' and HandoverToTendering eq 1"
          )
          .orderBy(
            'ExpectedClose',
            true
          )();

      /*
       * Convert SharePoint items
       * into execution projects.
       */

      const wonProjects: IExecutionProject[] =
        items.map(
          item => {

            const checklistValues =
              getChecklistValues(
                item.ExecutionChecklist
              );

            return {

              sharePointId:
                item.Id,

              id:
                item.OpportunityNumber ||
                String(item.Id),

              title:
                item.Title ||
                '',

              customer:
                item.Client ||
                '',

              domain:
                item.Category ||
                item.Team ||
                '',

              manager:
                item.TenderingManager ||
                '',

              engineer:
                item.TenderingEngineer ||
                '',

              deadline:
                item.ExpectedClose ||
                '',

              projectManager:
                item.ProjectManager ||
                '',

              checklist:
                createChecklist(
                  checklistValues
                )

            };

          }
        );

      setProjects(
        wonProjects
      );

    } catch (err) {

      console.error(
        'ERROR LOADING EXECUTION PROJECTS:',
        err
      );

      setError(
        'Unable to load execution projects from the Opportunities list.'
      );

    } finally {

      setLoading(false);

    }

  };

  /*
   * LOAD WHEN WEBPART OPENS
   */

  React.useEffect(
    () => {

      void loadProjects();

    },
    [context]
  );

  /*
   * CALCULATE PROGRESS
   */

  const calculateProgress = (
    checklist: IExecutionProject['checklist']
  ): number => {

    let completed = 0;

    checklistItems.forEach(
      item => {

        if (
          checklist[item.key]
        ) {

          completed++;

        }

      }
    );

    return Math.round(
      (
        completed /
        checklistItems.length
      ) * 100
    );

  };

  /*
   * SAVE PROJECT MANAGER
   */

  const assignProjectManager = async (
    projectId: string,
    manager: string
  ): Promise<void> => {

    const project =
      projects.find(
        item =>
          item.id === projectId
      );

    if (!project) {
      return;
    }

    try {

      setSaving(true);
      setError('');

      const sp = getSP(context);

      await sp.web.lists
        .getByTitle('Opportunities')
        .items
        .getById(
          project.sharePointId
        )
        .update({
          ProjectManager:
            manager
        });

      setProjects(
        currentProjects =>
          currentProjects.map(
            currentProject => {

              if (
                currentProject.id !==
                projectId
              ) {

                return currentProject;

              }

              return {

                ...currentProject,

                projectManager:
                  manager

              };

            }
          )
      );

    } catch (err) {

      console.error(
        'ERROR SAVING PROJECT MANAGER:',
        err
      );

      setError(
        'Unable to save the project manager.'
      );

    } finally {

      setSaving(false);

    }

  };

  /*
   * SAVE CHECKLIST
   */

  const toggleChecklistItem = async (
    projectId: string,
    key: keyof IExecutionProject['checklist']
  ): Promise<void> => {

    const project =
      projects.find(
        item =>
          item.id === projectId
      );

    if (!project) {
      return;
    }

    /*
     * Create new checklist state.
     */

    const updatedChecklist = {

      ...project.checklist,

      [key]:
        !project.checklist[key]

    };

    /*
     * Convert it to SharePoint choices.
     */

    const sharePointValues =
      checklistToSharePoint(
        updatedChecklist
      );

    const nextProgress =
      calculateProgress(
        updatedChecklist
      );

    try {

      setSaving(true);
      setError('');

      const sp = getSP(context);

      /*
       * IMPORTANT:
       *
       * ExecutionChecklist is a
       * SharePoint Choice column
       * with multiple selections.
       */

      await sp.web.lists
        .getByTitle('Opportunities')
        .items
        .getById(
          project.sharePointId
        )
        .update({

          ExecutionChecklist:
            sharePointValues,
          Stage:
            nextProgress === 100
              ? 'Completed'
              : 'In Execution'

        });

      /*
       * Update the screen only
       * after SharePoint succeeds.
       */

      setProjects(
        currentProjects =>
          currentProjects.map(
            currentProject => {

              if (
                currentProject.id !==
                projectId
              ) {

                return currentProject;

              }

              return {

                ...currentProject,

                checklist:
                  updatedChecklist

              };

            }
          )
      );

    } catch (err) {

      console.error(
        'ERROR SAVING EXECUTION CHECKLIST:',
        err
      );

      setError(
        'Unable to save the checklist.'
      );

    } finally {

      setSaving(false);

    }

  };

  /*
   * SEARCH
   */

  const filteredProjects =
    projects.filter(
      project => {

        const search =
          searchTerm
            .toLowerCase()
            .trim();

        if (!search) {
          return true;
        }

        return (

          project.id
            .toLowerCase()
            .indexOf(search) !== -1 ||

          project.title
            .toLowerCase()
            .indexOf(search) !== -1 ||

          project.customer
            .toLowerCase()
            .indexOf(search) !== -1 ||

          project.domain
            .toLowerCase()
            .indexOf(search) !== -1 ||

          project.manager
            .toLowerCase()
            .indexOf(search) !== -1 ||

          project.engineer
            .toLowerCase()
            .indexOf(search) !== -1 ||

          project.projectManager
            .toLowerCase()
            .indexOf(search) !== -1

        );

      }
    );

  /*
   * SORT
   */

  const sortedProjects =
    [...filteredProjects].sort(
      (a, b) => {

        switch (sortBy) {

          case 'deadlineAsc':

            return (
              new Date(
                a.deadline
              ).getTime() -
              new Date(
                b.deadline
              ).getTime()
            );

          case 'deadlineDesc':

            return (
              new Date(
                b.deadline
              ).getTime() -
              new Date(
                a.deadline
              ).getTime()
            );

          case 'nameAsc':

            return a.title.localeCompare(
              b.title
            );

          case 'nameDesc':

            return b.title.localeCompare(
              a.title
            );

          case 'progressAsc':

            return (
              calculateProgress(
                a.checklist
              ) -
              calculateProgress(
                b.checklist
              )
            );

          case 'progressDesc':

            return (
              calculateProgress(
                b.checklist
              ) -
              calculateProgress(
                a.checklist
              )
            );

          default:

            return 0;

        }

      }
    );

  /*
   * RENDER
   */

  return (

    <div className={styles.executionPage}>

      {/* HEADER */}

      <div className={styles.executionHeader}>

        <div>

          <div className={styles.eyebrow}>
            PROJECT EXECUTION
          </div>

          <h1>
            Execution Tracker
          </h1>

          <p>
            Manage projects that have been won
            and handed over from Tendering.
          </p>

        </div>

        <div className={styles.projectCount}>

          <strong>
            {projects.length}
          </strong>

          <span>
            Active Projects
          </span>

        </div>

      </div>

      {/* SAVING */}

      {saving && (

        <div
          style={{
            padding: '10px 15px',
            marginBottom: '15px',
            background: '#eff6ff',
            color: '#2563eb',
            borderRadius: '8px',
            fontSize: '13px'
          }}
        >
          Saving changes...
        </div>

      )}

      {/* ERROR */}

      {!loading &&
        error && (

          <div className={styles.emptyState}>

            <div
              className={
                styles.emptyStateIcon
              }
            >
              !
            </div>

            <h2>
              Unable to load projects
            </h2>

            <p>
              {error}
            </p>

          </div>

        )}

      {/* LOADING */}

      {loading && (

        <div className={styles.emptyState}>

          <div
            className={
              styles.emptyStateIcon
            }
          >
            ⏳
          </div>

          <h2>
            Loading execution projects...
          </h2>

          <p>
            Loading won opportunities from SharePoint.
          </p>

        </div>

      )}

      {/* SEARCH + SORT */}

      {!loading &&
        !error &&
        projects.length > 0 && (

          <div className={styles.toolbar}>

            <div className={styles.searchBox}>

              <span>
                🔍
              </span>

              <input
                type="text"
                placeholder="Search projects..."
                value={searchTerm}
                onChange={
                  event =>
                    setSearchTerm(
                      event.target.value
                    )
                }
              />

            </div>

            <div className={styles.sortBox}>

              <label>
                Sort by
              </label>

              <select
                value={sortBy}
                onChange={
                  event =>
                    setSortBy(
                      event.target.value
                    )
                }
              >

                <option value="default">
                  Default
                </option>

                <option value="deadlineAsc">
                  Deadline — Earliest First
                </option>

                <option value="deadlineDesc">
                  Deadline — Latest First
                </option>

                <option value="nameAsc">
                  Project Name — A to Z
                </option>

                <option value="nameDesc">
                  Project Name — Z to A
                </option>

                <option value="progressAsc">
                  Completion — Lowest First
                </option>

                <option value="progressDesc">
                  Completion — Highest First
                </option>

              </select>

            </div>

          </div>

        )}

      {/* NO PROJECTS */}

      {!loading &&
        !error &&
        projects.length === 0 && (

          <div className={styles.emptyState}>

            <div
              className={
                styles.emptyStateIcon
              }
            >
              ✓
            </div>

            <h2>
              No execution projects yet
            </h2>

            <p>
              When an opportunity is marked as Won,
              it will automatically appear here.
            </p>

          </div>

        )}

      {/* NO SEARCH RESULTS */}

      {!loading &&
        !error &&
        projects.length > 0 &&
        sortedProjects.length === 0 && (

          <div className={styles.emptyState}>

            <div
              className={
                styles.emptyStateIcon
              }
            >
              🔍
            </div>

            <h2>
              No projects found
            </h2>

            <p>
              Try changing your search.
            </p>

          </div>

        )}

      {/* PROJECT CARDS */}

      {!loading &&
        !error &&
        sortedProjects.length > 0 && (

          <div className={styles.projectGrid}>

            {sortedProjects.map(
              project => {

                const progress =
                  calculateProgress(
                    project.checklist
                  );

                const isEditing =
                  editingProjectId ===
                  project.id;

                return (

                  <article
                    key={
                      project.sharePointId
                    }
                    className={
                      progress === 100
                        ? `${styles.projectCard} ${styles.completedProject}`
                        : styles.projectCard
                    }
                  >

                    {/* CARD HEADER */}

                    <div
                      className={
                        styles.cardHeader
                      }
                    >

                      <div>

                        <div
                          className={
                            styles.projectId
                          }
                        >
                          {project.id}
                        </div>

                        <h2>
                          {project.title}
                        </h2>

                      </div>

                      <span
                        className={
                          progress === 100
                            ? `${styles.wonBadge} ${styles.completedBadge}`
                            : styles.wonBadge
                        }
                      >
                        {progress === 100
                          ? 'COMPLETED'
                          : 'WON'}
                      </span>

                    </div>

                    {/* INFORMATION */}

                    <div
                      className={
                        styles.infoGrid
                      }
                    >

                      <div
                        className={
                          styles.infoItem
                        }
                      >

                        <span>
                          CLIENT
                        </span>

                        <strong>
                          {project.customer}
                        </strong>

                      </div>

                      <div
                        className={
                          styles.infoItem
                        }
                      >

                        <span>
                          DOMAIN
                        </span>

                        <strong>
                          {project.domain}
                        </strong>

                      </div>

                      <div
                        className={
                          styles.infoItem
                        }
                      >

                        <span>
                          TENDERING MANAGER
                        </span>

                        <strong>
                          {project.manager}
                        </strong>

                      </div>

                      <div
                        className={
                          styles.infoItem
                        }
                      >

                        <span>
                          ENGINEER
                        </span>

                        <strong>
                          {project.engineer}
                        </strong>

                      </div>

                    </div>

                    {/* PROJECT MANAGER */}

                    <div
                      className={
                        styles.managerSection
                      }
                    >

                      <div>

                        <span
                          className={
                            styles.sectionLabel
                          }
                        >
                          PROJECT MANAGER
                        </span>

                        <strong
                          className={
                            styles.managerName
                          }
                        >
                          {project.projectManager ||
                            'Not assigned'}
                        </strong>

                      </div>

                      <select
                        className={
                          styles.managerSelect
                        }
                        value={
                          project.projectManager
                        }
                        disabled={
                          saving
                        }
                        onChange={
                          event =>
                            void assignProjectManager(
                              project.id,
                              event.target.value
                            )
                        }
                      >

                        <option value="">
                          Assign Manager
                        </option>

                        {projectManagers.map(
                          manager => (

                            <option
                              key={manager}
                              value={manager}
                            >
                              {manager}
                            </option>

                          )
                        )}

                      </select>

                    </div>

                    {/* PROGRESS */}

                    <div
                      className={
                        styles.progressSection
                      }
                    >

                      <div
                        className={
                          styles.progressHeader
                        }
                      >

                        <span>
                          EXECUTION COMPLETION
                        </span>

                        <strong>
                          {progress}%
                        </strong>

                      </div>

                      <div
                        className={
                          styles.progressBar
                        }
                      >

                        <div
                          className={
                            styles.progressFill
                          }
                          style={{
                            width:
                              `${progress}%`
                          }}
                        />

                      </div>

                      <div
                        className={
                          styles.progressText
                        }
                      >

                        {progress === 100

                          ? 'Execution completed'

                          : `${
                              checklistItems.filter(
                                item =>
                                  project.checklist[
                                    item.key
                                  ]
                              ).length
                            } of ${
                              checklistItems.length
                            } stages completed`
                        }

                      </div>

                    </div>

                    {/* EDIT BUTTON */}

                    <div
                      className={
                        styles.cardActions
                      }
                    >

                      <button
                        className={
                          styles.editButton
                        }
                        onClick={() =>
                          setEditingProjectId(
                            isEditing
                              ? null
                              : project.id
                          )
                        }
                      >

                        {isEditing
                          ? 'Close Checklist'
                          : 'Edit Handover Information & Checklist'}

                      </button>

                    </div>

                    {/* CHECKLIST */}

                    {isEditing && (

                      <div
                        className={
                          styles.checklistSection
                        }
                      >

                        <div
                          className={
                            styles.checklistHeader
                          }
                        >

                          <div>

                            <span
                              className={
                                styles.sectionLabel
                              }
                            >
                              EXECUTION CHECKLIST
                            </span>

                            <h3>
                              Project Stages
                            </h3>

                          </div>

                          <span
                            className={
                              styles.checklistPercentage
                            }
                          >
                            {progress}%
                          </span>

                        </div>

                        <div
                          className={
                            styles.checklist
                          }
                        >

                          {checklistItems.map(
                            item => (

                              <label
                                key={
                                  item.key
                                }
                                className={
                                  project.checklist[
                                    item.key
                                  ]
                                    ? `${styles.checklistItem} ${styles.completed}`
                                    : styles.checklistItem
                                }
                              >

                                <input
                                  type="checkbox"
                                  checked={
                                    project.checklist[
                                      item.key
                                    ]
                                  }
                                  disabled={
                                    saving
                                  }
                                  onChange={() =>
                                    void toggleChecklistItem(
                                      project.id,
                                      item.key
                                    )
                                  }
                                />

                                <span
                                  className={
                                    styles.checkboxCustom
                                  }
                                >
                                  {
                                    project.checklist[
                                      item.key
                                    ]
                                      ? '✓'
                                      : ''
                                  }
                                </span>

                                <span
                                  className={
                                    styles.checklistLabel
                                  }
                                >
                                  {item.label}
                                </span>

                                {project.checklist[
                                  item.key
                                ] && (

                                  <span
                                    className={
                                      styles.doneBadge
                                    }
                                  >
                                    DONE
                                  </span>

                                )}

                              </label>

                            )
                          )}

                        </div>

                      </div>

                    )}

                    {/* FOOTER */}

                    <div
                      className={
                        styles.cardFooter
                      }
                    >

                      <div>

                        <span>
                          DEADLINE
                        </span>

                        <strong>
                          {project.deadline ||
                            '-'}
                        </strong>

                      </div>

                      <div>

                        <span>
                          STATUS
                        </span>

                        <strong
                          className={
                            styles.executionStatus
                          }
                        >
                          {progress === 100
                            ? 'Completed'
                            : 'In Execution'}
                        </strong>

                      </div>

                    </div>

                  </article>

                );

              }
            )}

          </div>

        )}

    </div>

  );

};

export default Execution;