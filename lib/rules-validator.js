const fs = require("fs");

function loadSchema(schemaPath) {
  return JSON.parse(fs.readFileSync(schemaPath, "utf8"));
}

function formatPath(path) {
  return path || "<root>";
}

function typeOf(value) {
  if (Array.isArray(value)) return "array";
  if (value === null) return "null";
  return typeof value;
}

function validateAgainstSchema(value, schema, path = "") {
  const errors = [];
  const actualType = typeOf(value);

  if (schema.type && actualType !== schema.type) {
    errors.push(`${formatPath(path)} 类型应为 ${schema.type}，实际为 ${actualType}`);
    return errors;
  }

  if (schema.enum && !schema.enum.includes(value)) {
    errors.push(`${formatPath(path)} 必须是 ${schema.enum.join(", ")} 之一`);
  }

  if (schema.type === "string") {
    if (schema.minLength !== undefined && value.length < schema.minLength) {
      errors.push(`${formatPath(path)} 长度不能小于 ${schema.minLength}`);
    }
    if (schema.maxLength !== undefined && value.length > schema.maxLength) {
      errors.push(`${formatPath(path)} 长度不能大于 ${schema.maxLength}`);
    }
    if (schema.pattern && !(new RegExp(schema.pattern).test(value))) {
      errors.push(`${formatPath(path)} 格式不符合 ${schema.pattern}`);
    }
  }

  if (schema.type === "number") {
    if (!Number.isFinite(value)) {
      errors.push(`${formatPath(path)} 必须是有限数字`);
    }
    if (schema.minimum !== undefined && value < schema.minimum) {
      errors.push(`${formatPath(path)} 不能小于 ${schema.minimum}`);
    }
    if (schema.maximum !== undefined && value > schema.maximum) {
      errors.push(`${formatPath(path)} 不能大于 ${schema.maximum}`);
    }
  }

  if (schema.type === "object") {
    const properties = schema.properties || {};
    for (const key of schema.required || []) {
      if (!Object.prototype.hasOwnProperty.call(value, key)) {
        errors.push(`${formatPath(path ? `${path}.${key}` : key)} 为必填规则`);
      }
    }
    if (schema.additionalProperties === false) {
      for (const key of Object.keys(value)) {
        if (!Object.prototype.hasOwnProperty.call(properties, key)) {
          errors.push(`${formatPath(path ? `${path}.${key}` : key)} 不是允许的规则项`);
        }
      }
    }
    for (const [key, childSchema] of Object.entries(properties)) {
      if (Object.prototype.hasOwnProperty.call(value, key)) {
        errors.push(...validateAgainstSchema(value[key], childSchema, path ? `${path}.${key}` : key));
      }
    }
  }

  return errors;
}

function validateRules(rules, schema) {
  return validateAgainstSchema(rules, schema);
}

function assertValidRules(rules, schema, label = "规则") {
  const errors = validateRules(rules, schema);
  if (errors.length) {
    const error = new Error(`${label}校验失败`);
    error.statusCode = 400;
    error.details = errors;
    throw error;
  }
  return true;
}

module.exports = {
  assertValidRules,
  loadSchema,
  validateRules
};
