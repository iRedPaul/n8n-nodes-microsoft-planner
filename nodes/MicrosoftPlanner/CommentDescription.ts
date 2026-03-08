import { INodeProperties } from "n8n-workflow";

// ----------------------------------
//         Comment operations
// ----------------------------------
export const commentOperations: INodeProperties[] = [
  {
    displayName: "Operation",
    name: "operation",
    type: "options",
    noDataExpression: true,
    displayOptions: {
      show: {
        resource: ["comment"],
      },
    },
    options: [
      {
        name: "Create",
        value: "create",
        description: "Create a new comment on a task",
        action: "Create a comment",
      },
      {
        name: "Get Many",
        value: "getAll",
        description: 'Get many comments from a task',
        action: "Get many comments",
      },
    ],
    default: "create",
  },
];

// ----------------------------------
//         Comment fields
// ----------------------------------
export const commentFields: INodeProperties[] = [
  // ----------------------------------
  //         comment:create
  // ----------------------------------
  {
    displayName: "Task",
    name: "taskId",
    type: "resourceLocator",
    default: { mode: "id", value: "" },
    required: true,
    displayOptions: {
      show: {
        "@tool": [false],
        resource: ["comment"],
        operation: ["create"],
      },
    },
    modes: [
      {
        displayName: "By ID",
        name: "id",
        type: "string",
        placeholder: "e.g. rz1EH6N_a0aLpRm-2QifxZgAF5OL",
      },
    ],
    description: "The task to comment on",
  },
  {
    displayName: "Task ID",
    name: "taskId",
    type: "string",
    default: "",
    required: true,
    displayOptions: {
      show: {
        "@tool": [true],
        resource: ["comment"],
        operation: ["create"],
      },
    },
    description: "The Microsoft Planner task ID to comment on",
    placeholder: "e.g. rz1EH6N_a0aLpRm-2QifxZgAF5OL",
  },
  {
    displayName: "Content",
    name: "content",
    type: "string",
    typeOptions: {
      rows: 5,
    },
    required: true,
    displayOptions: {
      show: {
        resource: ["comment"],
        operation: ["create"],
      },
    },
    default: "",
    description: "The content of the comment",
  },
  {
    displayName: "Additional Fields",
    name: "additionalFields",
    type: "collection",
    placeholder: "Add Field",
    default: {},
    displayOptions: {
      show: {
        resource: ["comment"],
        operation: ["create"],
      },
    },
    options: [
      {
        displayName: "Content Type",
        name: "contentType",
        type: "options",
        displayOptions: {
          show: {
            "@tool": [false],
          },
        },
        options: [
          {
            name: "HTML",
            value: "html",
          },
          {
            name: "Text",
            value: "text",
          },
        ],
        default: "text",
        description: "The format of the comment content",
      },
      {
        displayName: "Content Type Text",
        name: "contentTypeText",
        type: "string",
        default: "",
        displayOptions: {
          show: {
            "@tool": [true],
          },
        },
        placeholder: "text",
        description:
          "Format of the comment content in tool mode. Supported values: text or html.",
      },
    ],
  },

  // ----------------------------------
  //         comment:getAll
  // ----------------------------------
  {
    displayName: "Task",
    name: "taskId",
    type: "resourceLocator",
    default: { mode: "id", value: "" },
    required: true,
    displayOptions: {
      show: {
        "@tool": [false],
        resource: ["comment"],
        operation: ["getAll"],
      },
    },
    modes: [
      {
        displayName: "By ID",
        name: "id",
        type: "string",
        placeholder: "e.g. rz1EH6N_a0aLpRm-2QifxZgAF5OL",
      },
    ],
    description: "The task to get comments from",
  },
  {
    displayName: "Task ID",
    name: "taskId",
    type: "string",
    default: "",
    required: true,
    displayOptions: {
      show: {
        "@tool": [true],
        resource: ["comment"],
        operation: ["getAll"],
      },
    },
    description:
      "The Microsoft Planner task ID whose comments should be returned",
    placeholder: "e.g. rz1EH6N_a0aLpRm-2QifxZgAF5OL",
  },
  {
    displayName: "Additional Fields",
    name: "additionalFields",
    type: "collection",
    placeholder: "Add Field",
    default: {},
    displayOptions: {
      show: {
        resource: ["comment"],
        operation: ["getAll"],
      },
    },
    options: [
      {
        displayName: "Select Properties",
        name: "select",
        type: "string",
        default: "",
        description: 'Comma-separated list of properties to specify which fields to return (e.g. ID,content)',
      },
    ],
  },
];
