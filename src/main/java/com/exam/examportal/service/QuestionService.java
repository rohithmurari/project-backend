package com.exam.examportal.service;

import java.util.Set;

import org.springframework.web.multipart.MultipartFile;

import com.exam.examportal.model.exam.Question;
import com.exam.examportal.model.exam.Quiz;

public interface QuestionService {
	public Question addQuestion(Question question);
	public Question updateQuestion(Question question);
	public Set<Question> getQuestions();
	public Question getQuestion(Long questionId);
	public Set<Question> getQuestionsOfQuiz(Quiz quiz);
	public void deleteQuestion(Long quesId);
	public Question get(Long questionsId);
	public void saveFromExcel(MultipartFile file);
	
}
