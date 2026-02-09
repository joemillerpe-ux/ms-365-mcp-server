import { z } from 'zod';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import GraphClient from './graph-client.js';
import logger from './logger.js';

interface PlannerTask {
  id: string;
  title: string;
  planId: string;
  bucketId?: string;
  priority: number;
  percentComplete: number;
  dueDateTime?: string;
  createdDateTime: string;
  assignments?: Record<string, unknown>;
}

interface PlannerTaskDetails {
  description?: string;
  checklist?: Record<string, { title: string; isChecked: boolean; orderHint?: string }>;
}

interface ChecklistItemResult {
  displayName: string;
  isChecked: boolean;
  success: boolean;
}

interface TodoTask {
  id: string;
  title: string;
  body?: { content: string; contentType: string };
  importance: string;
  status: string;
  dueDateTime?: { dateTime: string; timeZone: string };
}

interface TodoTaskList {
  id: string;
  displayName: string;
  wellknownListName: string;
}

/**
 * Maps Planner task data to To-Do task format
 */
function mapPlannerToTodo(
  task: PlannerTask,
  details: PlannerTaskDetails | null,
  addPrefix: boolean
): Record<string, unknown> {
  // Priority mapping: Planner (1=urgent, 5=normal, 9=low) -> To-Do (high, normal, low)
  const importanceMap: Record<number, string> = { 1: 'high', 5: 'normal', 9: 'low' };
  const title = addPrefix ? `[Planner] ${task.title}` : task.title;

  const todoTask: Record<string, unknown> = {
    title,
    body: {
      content: buildTaskBody(task, details),
      contentType: 'text',
    },
    importance: importanceMap[task.priority] || 'normal',
    status: task.percentComplete === 100 ? 'completed' : 'notStarted',
  };

  // Only add dueDateTime if present
  if (task.dueDateTime) {
    todoTask.dueDateTime = {
      dateTime: task.dueDateTime,
      timeZone: 'UTC',
    };
  }

  return todoTask;
}

/**
 * Builds the task body content with original description and metadata
 */
function buildTaskBody(task: PlannerTask, details: PlannerTaskDetails | null): string {
  const lines: string[] = [];

  // Add original description if present
  if (details?.description) {
    lines.push(details.description);
    lines.push('');
  }

  // Add conversion metadata
  lines.push('--- Converted from Planner ---');
  lines.push(`Plan ID: ${task.planId}`);
  lines.push(`Original Task ID: ${task.id}`);
  lines.push(`Created: ${task.createdDateTime}`);

  return lines.join('\n');
}

export function registerConversionTools(server: McpServer, graphClient: GraphClient): void {
  server.tool(
    'convert-planner-task-to-todo',
    'Converts a Planner task to a To-Do task. Useful for moving tasks from "Assigned to me" into your Tasks list for ROI tracking. By default, marks the Planner task complete and adds [Planner] prefix to the title. Keywords: move Planner to To-Do, convert Planner, migrate Planner task, transfer Planner',
    {
      plannerTaskId: z.string().describe('The ID of the Planner task to convert'),
      todoTaskListId: z
        .string()
        .optional()
        .describe('Target To-Do list ID (default: Tasks/defaultList)'),
      markPlannerComplete: z
        .boolean()
        .optional()
        .default(true)
        .describe('Mark the Planner task as complete after conversion (default: true)'),
      addPlannerPrefix: z
        .boolean()
        .optional()
        .default(true)
        .describe('Add [Planner] prefix to the task title (default: true)'),
    },
    async (params) => {
      try {
        logger.info(`Converting Planner task ${params.plannerTaskId} to To-Do`);

        // 1. Get the Planner task
        const plannerTaskResponse = await graphClient.graphRequest(
          `/planner/tasks/${params.plannerTaskId}`
        );
        const plannerTaskText =
          plannerTaskResponse.content[0].type === 'text' ? plannerTaskResponse.content[0].text : '';
        const plannerTask: PlannerTask = JSON.parse(plannerTaskText);

        // 2. Get task details (for description)
        let taskDetails: PlannerTaskDetails | null = null;
        try {
          const detailsResponse = await graphClient.graphRequest(
            `/planner/tasks/${params.plannerTaskId}/details`
          );
          const detailsText =
            detailsResponse.content[0].type === 'text' ? detailsResponse.content[0].text : '';
          taskDetails = JSON.parse(detailsText);
        } catch (error) {
          logger.warn(`Could not fetch task details: ${error}`);
          // Continue without details - not critical
        }

        // 3. Get the target To-Do list
        let listId = params.todoTaskListId;
        if (!listId) {
          const listsResponse = await graphClient.graphRequest('/me/todo/lists');
          const listsText =
            listsResponse.content[0].type === 'text' ? listsResponse.content[0].text : '';
          const listsData = JSON.parse(listsText);
          const tasksList = listsData.value?.find(
            (l: TodoTaskList) => l.wellknownListName === 'defaultList'
          );
          if (!tasksList) {
            throw new Error('Could not find default Tasks list');
          }
          listId = tasksList.id;
        }

        // 4. Map fields and create To-Do task
        const todoBody = mapPlannerToTodo(
          plannerTask,
          taskDetails,
          params.addPlannerPrefix !== false
        );

        const createResponse = await graphClient.graphRequest(`/me/todo/lists/${listId}/tasks`, {
          method: 'POST',
          body: JSON.stringify(todoBody),
        });
        const createText =
          createResponse.content[0].type === 'text' ? createResponse.content[0].text : '';
        const createdTodoTask: TodoTask = JSON.parse(createText);

        // 5. Create checklist items if present
        const checklistResults: ChecklistItemResult[] = [];
        if (taskDetails?.checklist) {
          // Sort by orderHint to maintain order (lower orderHint = higher in list)
          const sortedItems = Object.values(taskDetails.checklist).sort((a, b) => {
            const orderA = a.orderHint || '';
            const orderB = b.orderHint || '';
            return orderA.localeCompare(orderB);
          });

          for (const item of sortedItems) {
            try {
              await graphClient.graphRequest(
                `/me/todo/lists/${listId}/tasks/${createdTodoTask.id}/checklistItems`,
                {
                  method: 'POST',
                  body: JSON.stringify({
                    displayName: item.title,
                    isChecked: item.isChecked,
                  }),
                }
              );
              checklistResults.push({
                displayName: item.title,
                isChecked: item.isChecked,
                success: true,
              });
            } catch (error) {
              logger.warn(`Failed to create checklist item: ${item.title} - ${error}`);
              checklistResults.push({
                displayName: item.title,
                isChecked: item.isChecked,
                success: false,
              });
            }
          }
          logger.info(
            `Created ${checklistResults.filter((r) => r.success).length}/${sortedItems.length} checklist items`
          );
        }

        // 6. Handle Planner task disposition
        let plannerStatus = 'unchanged';
        if (params.markPlannerComplete !== false) {
          try {
            // Need to get ETag for Planner task update (required by Graph API)
            const taskWithEtag = await graphClient.graphRequest(
              `/planner/tasks/${params.plannerTaskId}`,
              { includeHeaders: true }
            );
            const taskWithEtagText =
              taskWithEtag.content[0].type === 'text' ? taskWithEtag.content[0].text : '';
            const taskData = JSON.parse(taskWithEtagText);
            const etag = taskData._etag || '*';

            await graphClient.graphRequest(`/planner/tasks/${params.plannerTaskId}`, {
              method: 'PATCH',
              body: JSON.stringify({ percentComplete: 100 }),
              headers: {
                'If-Match': etag,
              },
            });
            plannerStatus = 'completed';
          } catch (error) {
            logger.error(`Failed to mark Planner task complete: ${error}`);
            plannerStatus = 'failed_to_complete';
          }
        }

        const checklistCreated = checklistResults.filter((r) => r.success).length;
        const checklistTotal = checklistResults.length;

        const result = {
          success: true,
          message: `Converted Planner task to To-Do task`,
          todoTask: {
            id: createdTodoTask.id,
            title: createdTodoTask.title,
            listId: listId,
            checklistItems: checklistTotal > 0 ? `${checklistCreated}/${checklistTotal} created` : 'none',
          },
          plannerTask: {
            id: plannerTask.id,
            title: plannerTask.title,
            status: plannerStatus,
          },
        };

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      } catch (error) {
        logger.error(`Error converting Planner task: ${error}`);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                success: false,
                error: `Conversion failed: ${(error as Error).message}`,
              }),
            },
          ],
          isError: true,
        };
      }
    }
  );

  // =============================================================================
  // find-todo-task: Search for a task by title
  // =============================================================================
  server.tool(
    'find-todo-task',
    'Finds a To-Do task by title. If exact match found, returns full task details. If no exact match, returns lightweight list of partial matches for disambiguation. Searches ROI Tasks list by default.',
    {
      title: z.string().describe('Task title or search keywords (e.g., "#ENG #BuildPart - Tapped Hole")'),
      todoTaskListId: z
        .string()
        .optional()
        .describe('List ID to search (default: searches ROI Tasks, then other common lists)'),
      includeCompleted: z
        .boolean()
        .optional()
        .default(false)
        .describe('Include completed tasks in search (default: false)'),
    },
    async (params) => {
      try {
        logger.info(`Finding task with title: ${params.title}`);

        // Get list(s) to search
        const listsToSearch: string[] = [];

        if (params.todoTaskListId) {
          listsToSearch.push(params.todoTaskListId);
        } else {
          // Get all lists and prioritize ROI Tasks, Goal Phase Tasks, Ad-Hoc Quickies
          const listsResponse = await graphClient.graphRequest('/me/todo/lists');
          const listsText = listsResponse.content[0].type === 'text' ? listsResponse.content[0].text : '';
          const listsData = JSON.parse(listsText);

          const priorityOrder = ['ROI', 'Goal Phase', 'Ad-Hoc', 'Tasks'];
          const sortedLists = (listsData.value as TodoTaskList[]).sort((a, b) => {
            const aIndex = priorityOrder.findIndex(p => a.displayName.includes(p));
            const bIndex = priorityOrder.findIndex(p => b.displayName.includes(p));
            const aScore = aIndex === -1 ? 999 : aIndex;
            const bScore = bIndex === -1 ? 999 : bIndex;
            return aScore - bScore;
          });

          for (const list of sortedLists) {
            listsToSearch.push(list.id);
          }
        }

        // Extract a search keyword from the title (use first significant word after hashtags)
        const searchKeyword = params.title
          .replace(/#\w+/g, '') // Remove hashtags
          .replace(/[^\w\s]/g, ' ') // Remove special chars
          .trim()
          .split(/\s+/)
          .find(word => word.length > 3) || params.title.slice(0, 20);

        const statusFilter = params.includeCompleted ? '' : "status ne 'completed' and ";
        const filter = `${statusFilter}contains(title, '${searchKeyword.replace(/'/g, "''")}')`;

        interface FoundTask {
          task: Record<string, unknown>;
          listId: string;
          listName?: string;
        }

        const allMatches: FoundTask[] = [];
        let exactMatch: FoundTask | null = null;

        // Search each list until we find an exact match or exhaust all lists
        for (const listId of listsToSearch) {
          try {
            const response = await graphClient.graphRequest(
              `/me/todo/lists/${listId}/tasks?$filter=${encodeURIComponent(filter)}`
            );
            const responseText = response.content[0].type === 'text' ? response.content[0].text : '';
            const data = JSON.parse(responseText);

            if (data.value && Array.isArray(data.value)) {
              for (const task of data.value) {
                const foundTask: FoundTask = { task, listId };
                allMatches.push(foundTask);

                // Check for exact title match
                if (task.title === params.title) {
                  exactMatch = foundTask;
                  break;
                }
              }
            }

            if (exactMatch) break;
          } catch (error) {
            logger.warn(`Error searching list ${listId}: ${error}`);
            // Continue to next list
          }
        }

        // Return results
        if (exactMatch) {
          // Exact match found - return full task details
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  matchType: 'exact',
                  task: exactMatch.task,
                  listId: exactMatch.listId,
                }, null, 2),
              },
            ],
          };
        } else if (allMatches.length > 0) {
          // Partial matches - return lightweight list
          const lightweightFields = ['id', 'title', 'status', 'importance', 'dueDateTime'];
          const lightweightMatches = allMatches.map(m => {
            const light: Record<string, unknown> = { listId: m.listId };
            for (const field of lightweightFields) {
              if ((m.task as Record<string, unknown>)[field] !== undefined) {
                light[field] = (m.task as Record<string, unknown>)[field];
              }
            }
            return light;
          });

          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  matchType: 'partial',
                  message: `No exact match for "${params.title}". Found ${allMatches.length} partial match(es):`,
                  matches: lightweightMatches,
                }, null, 2),
              },
            ],
          };
        } else {
          // No matches
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  success: false,
                  matchType: 'none',
                  message: `No tasks found matching "${params.title}"`,
                  searchKeyword,
                }, null, 2),
              },
            ],
          };
        }
      } catch (error) {
        logger.error(`Error finding task: ${error}`);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                success: false,
                error: `Find failed: ${(error as Error).message}`,
              }),
            },
          ],
          isError: true,
        };
      }
    }
  );

  // =============================================================================
  // REMOVED: Attachment & LinkedResource Copy Tools (Feb 2026)
  // =============================================================================
  // Tools removed: copy-todo-task-attachments, copy-todo-linked-resources
  // Endpoints removed from endpoints.json:
  //   - list-todo-task-attachments, get-todo-task-attachment
  //   - create-todo-task-attachment, delete-todo-task-attachment
  //   - list-todo-linked-resources, get-todo-linked-resource
  //   - create-todo-linked-resource, delete-todo-linked-resource
  //
  // REASON: Microsoft Graph API does NOT expose ANY attachments on Outlook tasks.
  // The /attachments and /linkedResources endpoints return empty even when tasks
  // show attachments in To-Do/Outlook UI.
  //
  // TESTED:
  //   - Tasks with hasAttachments:true still return 0 attachments via API
  //   - Brand new tasks with dragged emails return empty
  //   - File attachments (Excel, PDF) dragged into tasks also return empty
  //   - expand=["attachments"] on list-todo-tasks returns no attachment data
  //   - linkedResources endpoint returns empty for all attachment types
  //   - The deprecated Outlook Tasks API (stopped Aug 2022) had itemAttachment
  //     support, but the current To-Do API does not expose this functionality.
  //
  // AFFECTS: ALL attachment types - emails, Excel, PDF, any file dragged into task
  //
  // WORKAROUND: Use Outlook Quick Step "Create a task with text of message"
  // instead of dragging emails. This embeds the email content in the task body,
  // making it accessible via the Graph API body property.
  //
  // See: C:\AI\McpServers\McpMs365Fork\docs\work-plan-attachments.md
  // =============================================================================
}
